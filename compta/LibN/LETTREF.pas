{***********UNITE*************************************************
Auteur  ...... : Compta
Créé le ...... : 01/01/1900
Modifié le ... : 31/07/2007
Description .. : 
Suite ........ : SBO 19/09/006 : FQ 18814 : Ajout du champ E_NOMLOT
Suite ........ : comme champ de référence. Attention modif de DFM
Suite ........ : 
Suite ........ : BVE 31.07.07 : Modification de la fonction de préparation 
Suite ........ : au lettrage pour gerer les gros volumes (3 fonctions ==> 1 
Suite ........ : seule)
Mots clefs ... : 
*****************************************************************}
unit LETTREF;

interface                     

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hcompte, HRegCpte, Hctrls, ComCtrls, hmsgbox, Menus, HSysMenu,
  HTB97, Ent1, HEnt1, Mask,HPanel, UiUtil, ExtCtrls,HStatus,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  TImpFic,ImpFicU,TabLiEdt,UtilEdt,
  LettUtil,SaisUtil, ed_Tools, LetBatch, UObjFiltres;

Type tLettRefAuto = Record
                    _Type : Integer ;
                    _Gen1,_Gen2,_Coll1,_Coll2,_Aux1,_Aux2,_CodeChamp,_ValChamp,_Exo,ChampMarque,_ValMarque : String ;
                    _Date1,_Date2 : tDateTime ;
                    _Partiel,_DebCre,_Montant,_MontantNeg,_Marque : Boolean ;
                    End ;

Type TTypeLettrage = (lettGene,lettAuxi,lettCegid) ;

Procedure LettrageParCode(OkSISCO : Boolean = FALSE ; OkCegid : Boolean = FALSE) ;
Function LettrageParCodeRV : Boolean ;
Function LettrageParCodePGISISCO : Boolean ;
Function LettrageParCodePGISISCOAuto(Var LRA : tLettRefAuto) : Boolean ;

type
  TFLET = class(TForm)
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    HPB: TToolWindow97;
    BValiderCEGID: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Msg: THMsgBox;
    Sauve: TSaveDialog;
    Pages: TPageControl;
    PStandards: TTabSheet;
    TCPTEDEBUT: THLabel;
    CPTEDEBUT: THCpteEdit;
    CPTEFIN: THCpteEdit;
    TCPTEFIN: THLabel;
    SauveIni: TSaveDialog;
    BFiltre: TToolbarButton97;
    TFExercice: THLabel;
    FExercice: THValComboBox;
    TFDateCpta1: THLabel;
    TFDateCpta2: TLabel;
    TFNomChamp: TLabel;
    TFCodeLettre1: TLabel;
    FCodeLettre1: TEdit;
    FTypeCompte: TRadioGroup;
    TFCodeLettre2: THLabel;
    FCodeLettre2: TEdit;
    FNomChamp: THValComboBox;
    FTraitePartiel: TCheckBox;
    FMontant: TCheckBox;
    FMontantNeg: TCheckBox;
    TCollDebut: THLabel;
    CollDebut: THCpteEdit;
    TCollFin: THLabel;
    CollFin: THCpteEdit;
    Hlabel2: THLabel;
    FNatCpt: THValComboBox;
    Timer1: TTimer;
    CBOkDEbCre: TCheckBox;
    TTravail: TLabel;
    TabSheet2: TTabSheet;
    FVol: TCheckBox;
    FMAJ: TCheckBox;
    F1: TLabel;
    TChampMAJ: TLabel;
    FChampMAJ: THValComboBox;
    TValMAJ: TLabel;
    FValMAJ: TEdit;
    F2: TLabel;
    FGroupLibres: TGroupBox;
    TFLibre1: THLabel;
    TFLibre2: THLabel;
    FLibre1: TEdit;
    FLibre2: TEdit;
    BDelet: TToolbarButton97;
    BOptim: TCheckBox;
    FDecoupeGen: TEdit;
    FDecoupeAux: TEdit;
    FDateCpta: THCritMaskEdit;
    FDateCpta_: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FTypeCompteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FTraitePartielClick(Sender: TObject);
    procedure FMontantClick(Sender: TObject);
    procedure FNatCptChange(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FMAJClick(Sender: TObject);
    procedure BValiderCEGIDClick(Sender: TObject);
    procedure FLibre1DblClick(Sender: TObject);
    procedure BDeletClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FDateCptaKeyPress(Sender: TObject; var Key: Char);
  private
    { Déclarations privées }
    NomFiltre       : String ;
    OkSISCO,OkCegid : Boolean ;
    PourRV,PbRV     : Boolean ;
    LRA             : tLettRefAuto ;
    ShunteInitMove  : Boolean ;
    ObjetFiltre     : TObjFiltre;
    F               : TextFile ;
    procedure LanceLettrage(lequel : TTypeLettrage);
    procedure TraiteLettrage( SQL : string ; lequel : TTypeLettrage ) ;
    function  GetSQLSegment( lequel : TTypeLettrage ) : string ;       
    function  GetSQLEcriture ( valeur : string ; lequel : TTypeLettrage) : string;
    procedure RecupMvt(Q : tQuery ; TLett : TList) ;
    (*
    procedure LanceLettrageJacadiTiers(Var GeneEnCours,AuxiEnCours : String) ;
    procedure LanceLettrageJacadiGene(Var GeneEnCours,AuxiEnCours : String) ;
    procedure LanceLettrageCEGIDTiers(Var GeneEnCours,AuxiEnCours : String) ; 
    Procedure QuelleBorne(What : Byte ; CptEnCours : String) ;*)
    procedure MajChamp(OnGen,OkPos : Boolean) ;
    Procedure UpdateLeLettrageCegid  ;
    procedure LitTRFParam ;
    Function  AlimChampDefaut : Boolean ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  paramdat;


Procedure LettrageParCode(OkSISCO : Boolean = FALSE ; OkCegid : Boolean = FALSE) ;
var FLET:TFLET ;
    PP : THPanel ;
BEGIN
FLET:=TFLET.Create(Application) ;
FLET.OkSISCO:=OkSISCO ;
FLET.OkCEGID:=OkCEGID ;
FLET.PourRV:=FALSE ;
FLET.PbRV:=FALSE ;
FLET.Timer1.Enabled:=FALSE ;
FLET.ShunteInitMove:=FALSE ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     FLET.ShowModal ;
     finally
     FLET.Free ;
     END ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FLET,PP) ;
   FLET.Show ;
   END ;
END ;

Function LettrageParCodeRV : Boolean ;
var FLET:TFLET ;
BEGIN
FLET:=TFLET.Create(Application) ;
FLET.OkSISCO:=FALSE ;
FLET.OkCEGID:=FALSE ;
FLET.PourRV:=TRUE ;
FLET.PbRV:=FALSE ;
FLET.ShunteInitMove:=FALSE ;
FLET.Timer1.Enabled:=TRUE ;
try
  FLET.ShowModal ;
  finally
  Result:=FLET.PbRV ;
  FLET.Free ;
  END ;
SourisNormale ;
END ;

Function LettrageParCodePGISISCO : Boolean ;
var FLET:TFLET ;
BEGIN
FLET:=TFLET.Create(Application) ;
FLET.OkSISCO:=TRUE ;
FLET.OkCEGID:=FALSE ;
FLET.PourRV:=TRUE ;
FLET.PbRV:=FALSE ;
FLET.ShunteInitMove:=FALSE ;
FLET.Timer1.Enabled:=TRUE ;
try
  FLET.ShowModal ;
  finally
  Result:=FLET.PbRV ;
  FLET.Free ;
  END ;
SourisNormale ;
END ;

Function LettrageParCodePGISISCOAuto(Var LRA : tLettRefAuto) : Boolean ;
var FLET:TFLET ;
BEGIN
FLET:=TFLET.Create(Application) ;
FLET.OkSISCO:=FALSE ;
FLET.OkCEGID:=FALSE ;
FLET.PourRV:=TRUE ;
FLET.PbRV:=FALSE ;
FLET.LRA:=LRA ;
FLET.ShunteInitMove:=TRUE ;
//FLET.Timer1.Enabled:=TRUE ;
try
  FLET.ShowModal ;
  finally
  Result:=FLET.PbRV ;
  FLET.Free ;
  END ;
SourisNormale ;
END ;



procedure DirDefault(Sauve : TSaveDialog ; FileName : String) ;
var j,i : integer ;
BEGIN
j:=Length(FileName);
for i:=Length(FileName) downto 1 do if FileName[i]='\' then BEGIN j:=i ; Break ; END ;
Sauve.InitialDir:=Copy(FileName,1,j) ;
END ;

procedure TFLET.LitTRFParam ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT TRP_DATA FROM TRFPARAM WHERE TRP_CODE="000" AND TRP_NOM="CP_TIERSL" ',FALSE) ;
If Not Q.Eof Then FDecoupeAux.Text:=Q.Fields[0].AsString ;
Ferme(Q) ;
Q:=OpenSQL('SELECT TRP_DATA FROM TRFPARAM WHERE TRP_CODE="000" AND TRP_NOM="CP_GENEL" ',FALSE) ;
If Not Q.Eof Then FDecoupeGen.Text:=Q.Fields[0].AsString ;
Ferme(Q) ;
END ;

Function TFLET.AlimChampDefaut : Boolean ;
BEGIN
Result:=TRUE ;
With LRA Do
  BEGIN
  If _CodeChamp='' Then BEGIN Result:=FALSE ; Exit ; END ;
  FTypeCompte.ItemIndex:=_Type  ;
  If _Type=0 Then
    BEGIN
    CPTEDEBUT.Text:=_Gen1 ;
    CPTEFIN.Text:=_Gen2 ;
    COLLDEBUT.Text:='' ;
    COLLFIN.Text:='' ;
    END Else
    BEGIN
    CPTEDEBUT.Text:=_Aux1 ;
    CPTEFIN.Text:=_Aux2 ;
    COLLDEBUT.Text:=_Coll1 ;
    COLLFIN.Text:=_Coll2 ;
    END ;
  FNomChamp.Value:=_CodeChamp ;
  If FNomChamp.ItemIndex<0 Then BEGIN Result:=FALSE ; Exit ; END ;
  FCodeLettre1.Text:=_ValChamp ;
  FCodeLettre2.Text:=_ValChamp ;
  If _Exo<>'' Then FExercice.Value:=_Exo Else FExercice.ItemIndex:=0 ;
  FDateCpta.Text:=DateToStr(_Date1) ;
  FDateCpta_.Text:=DateToStr(_Date2) ;
  FTraitePartiel.Checked:=_Partiel ;
  CBOkDebCre.Checked:=_DebCre ;
  FMontant.Checked:=_Montant ;
  FMontantNeg.Checked:=_MontantNeg ;
  FMaj.Checked:=_Marque ;
  if _Marque Then
    BEGIN
    FChampMaj.Value:=ChampMarque ;
    If FChampMaj.ItemIndex<0 Then FMaj.Checked:=FALSE ;
    FValMaj.Text:=_ValMarque ;
    If FValMaj.Text='' Then FMaj.Checked:=FALSE ;
    END ;
  END ;
END ;

procedure TFLET.FormShow(Sender: TObject);
Var OkAuto : Boolean ;
TCF        : TControlFiltre;
begin
OkAuto:=TRUE ;
FDecoupeGen.Text:='' ; FDecoupeAux.Text:='' ;
NomFiltre:='RECUPL' ;
FNomChamp.ItemIndex:=0 ;
Pages.ActivePage:=Pages.Pages[0] ;
If OkSISCO Then
  BEGIN
  FNomChamp.Value:='E_REFLIBRE' ;
  If VH^.RecupLTL Or VH^.RecupSISCOPGI Then FNomChamp.Value:='E_REFLETTRAGE' ;
  FTraitePartiel.Checked:=FALSE ;
  END Else If PourRV Then
  BEGIN
  If VH^.RecupSISCOPGI Then
    BEGIN
//    FNomChamp.Value:='E_REFLETTRAGE' ;
    OkAuto:=AlimChampDefaut ;
    END Else
    BEGIN
    FNomChamp.Value:='E_LIBRETEXTE0' ;
    FExercice.Value:=VH^.EnCours.Code ;
    FTraitePartiel.Checked:=TRUE ;
    END ;
  END ;
//If Not VH^.RecupSISCOPGI Then ChargeFiltre(NomFiltre,FFiltres,Pages) ;
If VH^.RecupSISCOPGI And (Not PourRV) Then
  BEGIN
  LitTRFParam ;
  END ;
FTypeCompte.SetFocus ;
//FNomChamp.Value:='E_LIBRETEXTE9' ;

// 01/07/2003 : Fiche 11734
FMontant.Visible := V_PGI.SAV;
// BPY le 09/09/2004 : fiche n° 10393
BOptim.Visible := V_PGI.SAV;

If OkCegid Then
  BEGIN
  BValider.Visible:=FALSE ; BValiderCEGID.Visible:=TRUE ;
  FTypeCompte.ItemIndex:=1 ; FTypeCompte.Enabled:=FALSE ;
  FNomChamp.Enabled:=FALSE ; TFNomChamp.Enabled:=FALSE ; FNomChamp.ItemIndex:=-1 ;
  FCodeLettre1.Enabled:=FALSE ; FCodeLettre2.Enabled:=FALSE ;
  TFCodeLettre1.Enabled:=FALSE ; TFCodeLettre2.Enabled:=FALSE ;
  FCodeLettre1.Text:='' ; FCodeLettre2.Text:='' ;
  FVol.Checked:=TRUE ; FMaj.Checked:=TRUE ; FChampMaj.Value:='E_LIBRETEXTE9' ; FValMAJ.Text:='LETTP' ;
  FMontant.Checked:=TRUE ;
  FTraitePartiel.Enabled:=FALSE ; CBOkDEbCre.Enabled:=FALSE ;
  FMontant.Enabled:=FALSE ; FMontantNeg.Enabled:=FALSE ;
  BDelet.Visible:=TRUE ;
  END Else
  BEGIN
  BValider.Visible:=TRUE ; BValiderCEGID.Visible:=FALSE ; BDelet.Visible:=FALSE ;
  END ;
Pages.ActivePage:=PStandards ;
If VH^.RecupSISCOPGI And PourRV And OkAuto Then Timer1.Enabled:=TRUE ;

// fiche 18816
TCF.Filtre := BFiltre;
TCF.Filtres := FFiltres;
TCF.PageCtrl := Pages;
ObjetFiltre := TObjFiltre.create(TCF, 'LETTREF');
ObjetFiltre.Charger;
end;


(*Procedure RecupMvtImportJacadi(Q : tQuery ; Var Mvt : TFMvtImport) ;
BEGIN
InitMvtImport(Mvt) ;
If Not Q.Eof Then With Mvt Do
  BEGIN
  IE_GENERAL:=Q.FindField('E_GENERAL').AsString ;
  IE_AUXILIAIRE:=Q.FindField('E_AUXILIAIRE').AsString ;
  IE_DATECOMPTABLE:=Q.FindField('E_DATECOMPTABLE').ASDateTime ;
  IE_DATEECHEANCE:=Q.FindField('E_DATEECHEANCE').ASDateTime ;
  IE_DATEREFEXTERNE:=Q.FindField('E_DATEREFEXTERNE').ASDateTime ;
  IE_DEVISE:=Q.FindField('E_DEVISE').AsString ;
  IE_LETTRAGE:=Q.FindField('E_LETTRAGE').AsString ;
  IE_JOURNAL:=Q.FindField('E_JOURNAL').AsString ;
  IE_DEBIT:=Q.FindField('E_DEBIT').AsFloat ;
  IE_CREDIT:=Q.FindField('E_CREDIT').ASFloat ;
  IE_REFINTERNE:=Q.FindField('E_REFINTERNE').AsString ;
  IE_REFEXTERNE:=Q.FindField('E_REFEXTERNE').AsString ; // Correction SBO
  IE_LIBELLE:=Q.FindField('E_LIBELLE').AsString ;
  IE_DEBITDEV:=Q.FindField('E_DEBITDEV').AsFloat ;
  IE_CREDITDEV:=Q.FindField('E_CREDITDEV').ASFloat ;
  IE_NATUREPIECE:=Q.FindField('E_NATUREPIECE').AsString ;
  IE_NUMLIGNE:=Q.FindField('E_NUMLIGNE').AsInteger ;
  IE_NUMECHE:=Q.FindField('E_NUMECHE').AsInteger ;
  IE_NUMPIECE:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  IE_TAUXDEV:=Q.FindField('E_TAUXDEV').ASFloat ;
  END
END ;          *)


Procedure PrepareLettrageJacadi(TLett : TList ; Var Mvt : TFMvtImport) ;
Var
    L : TL_Rappro ;
BEGIN
L:=TL_Rappro.Create ;
L.General:=Mvt.IE_GENERAL ; L.Auxiliaire:=Mvt.IE_AUXILIAIRE ;
L.DateC:=Mvt.IE_DATECOMPTABLE ; L.DateE:=Mvt.IE_DATEECHEANCE ; L.DateR:=Mvt.IE_DATEREFEXTERNE ;
L.RefI:=Mvt.IE_REFINTERNE ; L.RefL:=Mvt.IE_REFLIBRE ;
L.RefE:=Mvt.IE_REFEXTERNE; L.Lib:=Mvt.IE_LIBELLE ;
L.Jal:=Mvt.IE_JOURNAL ; L.Numero:=Mvt.IE_NUMPIECE ;
L.NumLigne:=Mvt.IE_NUMLIGNE ; L.NumEche:=Mvt.IE_NUMECHE ;
L.CodeL:=Mvt.IE_LETTRAGE ;
L.TauxDEV:=Mvt.IE_TAUXDEV ;
L.CodeD:=Mvt.IE_DEVISE ;
L.Decim:=V_PGI.OkDecV ;
L.Debit:=Mvt.IE_DEBIT ; L.Credit:=Mvt.IE_CREDIT ;
L.DebitCur:=Mvt.IE_DEBIT ; L.CreditCur:=Mvt.IE_CREDIT ;
L.DebDev:=Mvt.IE_DEBITDEV ; L.CredDev:=Mvt.IE_CREDITDEV ;
L.Nature:=Mvt.IE_NATUREPIECE ;
L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
L.Solution:=0 ; L.Exo:=QUELEXODT(Mvt.IE_DATECOMPTABLE) ;
L.EditeEtatTva:=Mvt.IE_LIBREBOOL1='X' ; 
TLett.Add(L) ;
END ;

Procedure FORCECOMMIT(Var Nb : Integer) ;
BEGIN
Inc(Nb) ;
If Nb>100 Then
  BEGIN
  CommitTrans ; BeginTrans ; Nb:=0 ;
  END ;
END ;

procedure TFLET.MajChamp(OnGen,OkPos : Boolean) ;
Var St : String ;
    Cpte1,Cpte2,Coll1,Coll2 : String ;
    NomChamp : String ;
    Date1,Date2 : TDateTime ;
    StV8 : String ;
BEGIN
NomChamp:=FNomChamp.Value ;
Cpte1:=CpteDebut.Text ; Cpte2:=CpteFin.Text ;
Coll1:=CollDebut.Text ; Coll2:=CollFin.Text ;
Date1:=StrToDate(FDateCpta.Text) ; Date2:=StrToDate(FDateCpta_.Text) ;
St:='UPDATE ECRITURE SET '+NomChamp+'=' ;
If OkPos Then St:=St+'E_DEBIT+E_CREDIT WHERE '
         Else St:=St+'(E_DEBIT+E_CREDIT)*(-1) WHERE ' ;
St:=St+' E_DATECOMPTABLE>="'+USDATETIME(Date1)+'" ' ;
St:=St+' And E_DATECOMPTABLE<="'+USDATETIME(Date2)+'" ' ;
If OnGen Then
  BEGIN
  If Cpte1<>'' Then St:=St+' AND E_GENERAL>="'+Cpte1+'" ' ;
  If Cpte2<>'' Then St:=St+' AND E_GENERAL<="'+Cpte2+'" ' ;
  END Else
  BEGIN
  If Cpte1<>'' Then St:=St+' AND E_AUXILIAIRE>="'+Cpte1+'" ' ;
  If Cpte2<>'' Then St:=St+' AND E_AUXILIAIRE<="'+Cpte2+'" ' ;
  If Coll1<>'' Then St:=St+' AND E_GENERAL>="'+Coll1+'" ' ;
  If Coll2<>'' Then St:=St+' AND E_GENERAL<="'+Coll2+'" ' ;
  END ;
If FExercice.ITemIndex>0 Then St:=St+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
// Correction SBO 11/08/2004 : Date testée pour ano typé H fausse !
St:=St+' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" Or E_ECRANOUVEAU="H") ' ;
// Fin Correction SBO
StV8:=LWhereV8 ; if StV8<>'' then St:=St+' AND '+StV8+' ' ;
If OkPos Then St:=St+' AND E_DEBIT+E_CREDIT>0 ' Else St:=St+' AND E_DEBIT+E_CREDIT<0 ' ;
Try
 BeginTrans ;
 ExecuteSQL(St) ;
 CommitTrans ;
Except
 Rollback ;
 End ;
END ;

Procedure VerifDevise(TLett : TList ; DevCur,SaisieCur : String) ;
Var What : integer ;
    i : Integer ;
    TL : TL_Rappro ;
BEGIN
What:=0 ;
If (DevCur='PIVOT') Or (SaisieCur='PIVOT') Then Exit ;
If (DevCur=V_PGI.DevisePivot) And (SaisieCur='-') Then Exit ;
If (DevCur=V_PGI.DevisePivot) And (SaisieCur='X') Then What:=2 ;
If (DevCur<>V_PGI.DevisePivot) Then What:=1 ;
for i:=0 to TLett.Count-1 do
    BEGIN
    TL:=TL_Rappro(TLett[i]) ;
    (*
    XD:=TL.DebDev ; XC:=TL.CredDev ;
    TL.DebDev:=TL.Debit ; TL.CredDev:=TL.Credit ;
    TL.Debit:=XD ; TL.Credit:=XC ;
    *)
    if What=1 then
       BEGIN
       TL.DebitCur:=TL.DebDev ; TL.CreditCur:=TL.CredDev ;
       END else if What=2 then
       BEGIN
       //TL.DebitCur:=TL.DebEuro ; TL.CreditCur:=TL.CredEuro ;
       END ;
    END ;
END ;

Procedure ouvreFicRapport(Var F : TextFile ; Cpt : String ; OnGen : Boolean) ;
Var LaDir : String ;
    StCR : String ;
BEGIN
{$i-}
LaDir:=ExtractFilePath(Application.ExeName) ;
If OnGen Then AssignFile(F,LaDir+'CRLETTRAGEGen.TXT')
         Else AssignFile(F,LaDir+'CRLETTRAGEAUX.TXT') ;
If Cpt='' Then Rewrite(F) Else Append(F) ;
If Cpt='' Then
  BEGIN
  StCR:='' ; Writeln(F,StCR) ;
  If OnGen Then STCR:='                COMPTE RENDU DE LETTRAGE DES GENERAUX'
           Else STCR:='                COMPTE RENDU DE LETTRAGE DES TIERS' ;
  Writeln(F,StCR) ;
  StCR:='' ; Writeln(F,StCR) ;
  StCR:=' * '+Format_String(' ',35)+' * '+Format_String(' ',4)+' * '+Format_String(' ',5)+'                *' ;
//  StCR:=' * '+Format_String(Gene,35)+' * '+CodeL+' * '+FormatFloat('00000',NbLettImp)+' Lignes Lettrées *' ;
  Writeln(F,StCR) ;
  END ;

{$i+}
END ;

Procedure FinFicRapport(Var F : TextFile) ;
Var StCR : String ;
BEGIN
{$i-}
StCR:=' * '+Format_String(' ',35)+' * '+Format_String(' ',4)+' * '+Format_String(' ',5)+'               *' ;
//  StCR:=' * '+Format_String(Gene,35)+' * '+CodeL+' * '+FormatFloat('00000',NbLettImp)+' Lignes Lettrées *' ;
{$i-}
Writeln(F,StCR) ;
{$i+}
END ;

Procedure EcritFicRapport(Var F : TextFile ; Cpt,CodeL : String ; NbLettImp : Integer) ;
Var StCR : String ;
BEGIN
StCR:=' * '+Format_String(Cpt,35)+' * '+CodeL+' * '+FormatFloat('00000',NbLettImp)+' Lignes Lettrées *' ;
{$i-}
Writeln(F,StCR) ;
{$i+}
END ;

(*Procedure TFLET.QuelleBorne(What : Byte ; CptEnCours : String) ;
var St,St1 : String ;
    fb : tFichierBase ;
    OkOk : Boolean ;
BEGIN
Exit ;
OkOk := false;
If Not VH^.RecupSISCOPGI Then Exit ;
Case What Of
  0 : BEGIN St:=FDecoupeGen.Text ; fb:=fbGene ; END ;
  1 : BEGIN St:=FDecoupeAux.Text ; fb:=fbAux ; END ;
  Else Exit ;
  END ;
If (St='') Then Exit ;
If CptEnCours='' Then
  BEGIN
  OkOk:=TRUE ;
  St1:=ReadTokenSt(St) ; St1:=BourreLaDonc(St1,fb) ;
  END Else
  BEGIN
  If (CpteFin.Text<>'') And (CpteFin.Text>CptEnCours) Then Exit ;
  While St1<CptEnCours Do
    BEGIN
    If St='' Then OkOk:=FALSE Else
      BEGIN
      St1:=ReadTokenSt(St) ; St1:=BourreLaDonc(St1,fb) ;
      If St1>CptEnCours Then BEGIN OkOk:=TRUE ; Break ; END ;
      END ;
    END ;
  END ;
If OkOk Then CpteFin.Text:=St1 Else CpteFin.Text:='' ;
END ;

procedure TFLET.LanceLettrageJacadiTiers(Var GeneEnCours,AuxiEnCours : String) ;
var SQL : String ;
    QImp : TQuery ;
    TLett:TList ;
    OkALettrer : Boolean ;
    NbLettImp : integer ;
    Auxi,Gene,Affaire,Devise : String ;
    MvtImport : FMvtImport ;
    TotD,TotC : Double ;
    Cpte1,Cpte2,Coll1,Coll2 : String ;
    NbTrans : Integer ;
    NomChamp : String ;
    Date1,Date2 : TDateTime ;
    StV8 : String ;
    DevCur,SaisieCur : String ;
    OnStoppe : Boolean ;
    NbLig,NbCpt : Integer ;
    F : TextFile ;
    CodeL : String ;
    ChampMaj,ValMaj : String ;
    SQLWhereTL : String ;
    Lefb : tfichierBase ;
BEGIN
  New(MvtImport) ;
  OuvreFicRapport(F,AuxiEnCours,FALSE) ;
  NomChamp:=FNomChamp.Value ;
  OnStoppe:=FALSE ;
  NbLig:=0 ;
  NbCpt:=0 ;
  QuelleBorne(1,AuxiEnCours) ;
  Cpte1:=CpteDebut.Text ;
  Cpte2:=CpteFin.Text ;
  NbTrans:=0 ;
  Coll1:=CollDebut.Text ;
  Coll2:=CollFin.Text ;
  Date1:=StrToDate(FDateCpta1.Text) ;
  Date2:=StrToDate(FDateCpta2.Text) ;
  SQL:='SELECT E_AUXILIAIRE, E_GENERAL, E_DEVISE, E_REFEXTERNE, E_REFLIBRE, E_AFFAIRE';
  SQL:=SQL+', E_LIBRETEXTE0, E_LIBRETEXTE1, E_LIBRETEXTE2, E_LIBRETEXTE3, E_LIBRETEXTE4, E_LIBRETEXTE5';
  SQL:=SQL+', E_LIBRETEXTE6, E_LIBRETEXTE7, E_LIBRETEXTE8, E_LIBRETEXTE9, E_REFRELEVE, E_REFLETTRAGE, E_NOMLOT';
  SQL:=SQL+', E_DATECOMPTABLE, E_DATEECHEANCE, E_DATEREFEXTERNE, E_LETTRAGE, E_JOURNAL, E_DEBIT, E_CREDIT';
  SQL:=SQL+', E_REFINTERNE, E_LIBELLE, E_DEBITDEV, E_CREDITDEV, E_NATUREPIECE, E_NUMLIGNE, E_NUMECHE';
  SQL:=SQL+', E_NUMEROPIECE, E_TAUXDEV' ;
  SQL:=SQL+' FROM ECRITURE ' ;
  SQL:=SQL+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
  SQL:=SQL+'WHERE '+NomChamp+'<>"" AND '
       +'E_AUXILIAIRE<>"" AND E_LETTRAGE="" AND E_ETATLETTRAGE="AL" ' ;
  IF FCodeLettre1.Text<>'' Then SQL:=SQL+' AND '+NomChamp+'>="'+FCodeLettre1.Text+'" ' ;
  IF FCodeLettre2.Text<>'' Then SQL:=SQL+' AND '+NomChamp+'<="'+FCodeLettre2.Text+'" ' ;
  If Cpte1<>'' Then SQL:=SQL+' AND E_AUXILIAIRE>="'+Cpte1+'" ' ;
  If Cpte2<>'' Then SQL:=SQL+' AND E_AUXILIAIRE<="'+Cpte2+'" ' ;
  If Coll1<>'' Then SQL:=SQL+' AND E_GENERAL>="'+Coll1+'" ' ;
  If Coll2<>'' Then SQL:=SQL+' AND E_GENERAL<="'+Coll2+'" ' ;
  If AuxiEnCours<>'' Then SQL:=SQL+' AND E_AUXILIAIRE>="'+AuxiEnCours+'" ' ;
  If GeneEnCours<>'' Then SQL:=SQL+' AND E_GENERAL>="'+GeneEnCours+'" ' ;
  If FNatCpt.ItemIndex>0 Then SQL:=SQL+' AND T_NATUREAUXI="'+FNatCpt.Value+'" ' ;
  SQL:=SQL+' And E_DATECOMPTABLE>="'+USDATETIME(Date1)+'" ' ;
  SQL:=SQL+' And E_DATECOMPTABLE<="'+USDATETIME(Date2)+'" ' ;
  If FExercice.ITemIndex>0 Then SQL:=SQL+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
  // Correction SBO 11/08/2004 : Date testée pour ano typé H fausse !
  SQL:=SQL+' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" Or E_ECRANOUVEAU="H") ' ;
  // Fin correction SBO
  StV8:=LWhereV8 ;
  if StV8<>'' then SQL:=SQL+' AND '+StV8+' ' ;

  if FTypeCOmpte.ItemIndex=0 then
     LeFb:=fbGene
  Else
     LeFb:=fbAux ;
  SQLWhereTL:=WhereLibre(FLibre1.Text,FLibre2.Text,LeFb,TRUE) ;
  If SQLWhereTL<>'' Then SQL:=SQL+SQLWHERETL ;

  SQL:=SQL+' ORDER BY E_AUXILIAIRE,E_GENERAL,E_DEVISE, '+NomChamp+', E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ' ;
  QImp:=OpenSQL(SQL,True) ;
  if (QImp.EOF) then
  begin
     Ferme(QImp) ;
     Exit ;
  end;
  If Not ShunteInitMove Then InitMove(RecordsCount(QImp),'') ;
  TLett:=TList.Create ;
  DevCur:='' ;
  SaisieCur:='' ;
  While not QImp.Eof do
  begin
     VideListe(TLett);
     NbLettImp:=0;
     TotD:=0;
     TotC:=0;
     RecupMvtImportJacadi(Qimp,MvtImport^) ;
     Auxi:=QImp.FindField('E_AUXILIAIRE').AsString ;
     Gene:=QImp.FindField('E_GENERAL').AsString ;
     Affaire:=QImp.FindField(NomChamp).AsString ;
     Devise:=QImp.FindField('E_DEVISE').AsString ;
     OkALettrer:=True ;
     While (not QImp.Eof) and ((QImp.FindField('E_AUXILIAIRE').AsString=Auxi)
       and (QImp.FindField('E_GENERAL').AsString=Gene)
       and (QImp.FindField(NomChamp).AsString=Affaire)
       and (QImp.FindField('E_DEVISE').AsString=Devise)) do
     begin
        Inc(NbLettImp) ;
        RecupMvtImportJacadi(QImp,MvtImport^) ;
        TotD:=Arrondi(TotD+MvtImport^.IE_DEBIT,V_PGI.OkDecV) ;
        TotC:=Arrondi(TotC+MvtImport^.IE_CREDIT,V_PGI.OkDecV) ;
        PrepareLettrageJacadi(TLett,MvtImport^) ;
        If (DevCur<>'PIVOT') And (SaisieCur<>'PIVOT') Then
        begin
           If DevCur='' Then DevCur:=MvtImport^.IE_DEVISE ;
           If DevCur<>MvtImport^.IE_DEVISE Then DevCur:='PIVOT' ;
           SaisieCur:='PIVOT' ;
        end ;
        QImp.Next ;
        If Not ShunteInitMove Then MoveCur(False) ;
        If FMontant.Checked And (Arrondi(TotD-TotC,V_PGI.OkDecV)=0) Then Break ;
     end;
     OkALettrer:=(OkALettrer and ((Arrondi(TotD-TotC,V_PGI.OkDecV)=0) Or FTraitePartiel.Checked)) And (NbLettImp>1) ;
     If CBOkDebCre.Checked Then
        OkALettrer:=OkALettrer And ((Arrondi(TotD,V_PGI.OkDecV)<>0) And (Arrondi(TotC,V_PGI.OkDecV)<>0)) ;
     if OkALettrer then
     begin
        ChampMaj:='' ;
        ValMaj:='' ;
        If FMaj.Checked And (FChampMaj.Value<>'') And (Trim(FValMaj.text)<>'') Then
        begin
           ChampMaj:=FChampMaj.Value ;
           ValMaj:=FValMaj.Text ;
        end;
        CodeL :=LettrerUnPaquet(TLett,False,False,ChampMaj,ValMaj) ;
        TTravail.Caption:='Lettrage '+Gene+' '+Auxi+' '+CodeL ;
        Application.ProcessMessages ;
        EcritFicRapport(F,Gene+' '+Auxi,CodeL,NbLettImp) ;
        DevCur:='' ;
        SaisieCur:='' ;
        Inc(NbCpt) ;
        NbLig:=NbLig+NbLettImp ;
     end;
     ForceCommit(NbTrans) ;
     OnStoppe:=(NbLig>5000) Or (NbCpt>20) ;
     If Not FVol.Checked Then OnStoppe:=FALSE ;
     If OnStoppe Then
     begin
        GeneEnCours:=Gene ;
        AuxiEnCours:=Auxi ;
        Break ;
     end;
  end;
  If Not OnStoppe Then
  begin
     GeneEnCours:='' ;
     AuxiEnCours:='' ;
     FinFicRapport(F) ;
  end;
  {$i-}
  CloseFile(F) ;
  {$i+}
  If Not ShunteInitMove Then FiniMove ;
  VideListe(TLett) ; TLett.Free ;
  Ferme(QImp) ;
  Dispose(MvtImport) ;
END ;

procedure TFLET.LanceLettrageJacadiGene(Var GeneEnCours,AuxiEnCours : String) ;
var SQL : String ;
    QImp : TQuery ;
    TLett:TList ;
    OkALettrer : Boolean ;
    NbLettImp : integer ;
    Auxi,Gene,Affaire,Devise : String ;
    MvtImport : FMvtImport ;
    TotD,TotC : Double ;
    Cpte1,Cpte2 : String ;
    NbTrans : Integer ;
    NomChamp : String ;
    Date1,Date2 : TDateTime ;
    StV8 : String ;
    OnStoppe : Boolean ;
    NbLig,NbCpt : Integer ;
    F : TextFile ;
    CodeL : String ;
    ChampMaj,ValMaj : String ;
    SQLWhereTL : String ;
    Lefb : tfichierBase ;
BEGIN
  New(MvtImport) ;
  NbTrans:=0 ;
  OnStoppe:=FALSE ;
  NbLig:=0 ;
  NbCpt:=0 ;
  OuvreFicRapport(F,AuxiEnCours,TRUE) ;
  NomChamp:=FNomChamp.Value ;
  Cpte1:=CpteDebut.Text ;
  Cpte2:=CpteFin.Text ;
  Date1:=StrToDate(FDateCpta1.Text) ;
  Date2:=StrToDate(FDateCpta2.Text) ;
  SQL:='SELECT E_AUXILIAIRE, E_GENERAL, E_DEVISE, E_REFEXTERNE, E_REFLIBRE, E_AFFAIRE';
  SQL:=SQL+', E_LIBRETEXTE0, E_LIBRETEXTE1, E_LIBRETEXTE2, E_LIBRETEXTE3, E_LIBRETEXTE4, E_LIBRETEXTE5';
  SQL:=SQL+', E_LIBRETEXTE6, E_LIBRETEXTE7, E_LIBRETEXTE8, E_LIBRETEXTE9, E_REFRELEVE, E_REFLETTRAGE, E_NOMLOT';
  SQL:=SQL+', E_DATECOMPTABLE, E_DATEECHEANCE, E_DATEREFEXTERNE, E_LETTRAGE, E_JOURNAL, E_DEBIT, E_CREDIT';
  SQL:=SQL+', E_REFINTERNE, E_LIBELLE, E_DEBITDEV, E_CREDITDEV, E_NATUREPIECE, E_NUMLIGNE, E_NUMECHE';
  SQL:=SQL+', E_NUMEROPIECE, E_TAUXDEV' ;
  SQL:=SQL+' FROM ECRITURE ' ;
  SQL:=SQL+' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
  SQL:=SQL+'WHERE '+NomChamp+'<>"" AND '
       +'E_AUXILIAIRE="" AND E_LETTRAGE="" AND E_ETATLETTRAGE="AL" ' ;
  IF FCodeLettre1.Text<>'' Then SQL:=SQL+' AND '+NomChamp+'>="'+FCodeLettre1.Text+'" ' ;
  IF FCodeLettre2.Text<>'' Then SQL:=SQL+' AND '+NomChamp+'<="'+FCodeLettre2.Text+'" ' ;
  If Cpte1<>'' Then SQL:=SQL+' AND E_GENERAL>="'+Cpte1+'" ' ;
  If Cpte2<>'' Then SQL:=SQL+' AND E_GENERAL<="'+Cpte2+'" ' ;
  If GeneEnCours<>'' Then SQL:=SQL+' AND E_GENERAL>="'+GeneEnCours+'" ' ;
  If FNatCpt.ItemIndex>0 Then SQL:=SQL+' AND G_NATUREGENE="'+FNatCpt.Value+'" ' ;
  SQL:=SQL+' And E_DATECOMPTABLE>="'+USDATETIME(Date1)+'" ' ;
  SQL:=SQL+' And E_DATECOMPTABLE<="'+USDATETIME(Date2)+'" ' ;
  If FExercice.ITemIndex>0 Then SQL:=SQL+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
  // Correction SBO 11/08/2004 : Date testée pour ano typé H fausse !
  SQL:=SQL+' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" Or E_ECRANOUVEAU="H") ' ;
  // Fin Correction SBO
  StV8:=LWhereV8 ;
  if StV8<>'' then SQL:=SQL+' AND '+StV8+' ' ;

  if FTypeCOmpte.ItemIndex=0 then
     LeFb:=fbGene
  else
     LeFb:=fbAux ;
  SQLWhereTL:=WhereLibre(FLibre1.Text,FLibre2.Text,LeFb,TRUE) ;
  If SQLWhereTL<>'' Then SQL:=SQL+SQLWHERETL ;
  SQL:=SQL+' ORDER BY E_AUXILIAIRE,E_GENERAL,E_DEVISE,'+NomChamp+', E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ' ;
  QImp:=OpenSQL(SQL,True) ;
  if (QImp.EOF) then
  begin
     Ferme(QImp) ;
     Exit ;
  end;
  If Not ShunteInitMove Then InitMove(RecordsCount(QImp),'') ;
  TLett:=TList.Create ;
  While not QImp.Eof do
  begin
     VideListe(TLett) ;
     NbLettImp:=0 ;
     TotD:=0 ;
     TotC:=0 ;
     RecupMvtImportJacadi(Qimp,MvtImport^) ;
     Auxi:=QImp.FindField('E_AUXILIAIRE').AsString ;
     Gene:=QImp.FindField('E_GENERAL').AsString ;
     Affaire:=QImp.FindField(NomChamp).AsString ;
     Devise:=QImp.FindField('E_DEVISE').AsString ;
     OkALettrer:=True ;
     While (not QImp.Eof) and ((QImp.FindField('E_AUXILIAIRE').AsString=Auxi)
       and (QImp.FindField('E_GENERAL').AsString=Gene)
       and (QImp.FindField(NomChamp).AsString=Affaire)
       and (QImp.FindField('E_DEVISE').AsString=Devise)) do
     begin
        Inc(NbLettImp) ;
        RecupMvtImportJacadi(QImp,MvtImport^) ;
        TotD:=Arrondi(TotD+MvtImport^.IE_DEBIT,V_PGI.OkDecV) ;
        TotC:=Arrondi(TotC+MvtImport^.IE_CREDIT,V_PGI.OkDecV) ;
        PrepareLettrageJacadi(TLett,MvtImport^) ;
        QImp.Next ;
        If Not ShunteInitMove Then MoveCur(False) ;
     end ;
     OkALettrer:=(OkALettrer and ((Arrondi(TotD-TotC,V_PGI.OkDecV)=0) Or FTraitePartiel.Checked)) And (NbLettImp>1) ;
     If CBOkDebCre.Checked Then
        OkALettrer:=OkALettrer And ((Arrondi(TotD,V_PGI.OkDecV)<>0) And (Arrondi(TotC,V_PGI.OkDecV)<>0)) ;
     if OkALettrer then
     begin
        ChampMaj:='' ; ValMaj:='' ;
        If FMaj.Checked And (FChampMaj.Value<>'') And (Trim(FValMaj.text)<>'') Then
        begin
           ChampMaj:=FChampMaj.Value ; ValMaj:=FValMaj.Text ;
        end;
        CodeL:=LettrerUnPaquet(TLett,False,False,ChampMaj,ValMaj) ;
        TTravail.Caption:='Lettrage '+Gene+' '+Auxi+' '+CodeL ; Application.ProcessMessages ;
        EcritFicRapport(F,Gene,CodeL,NbLettImp) ;
        Inc(NbCpt) ;
        NbLig:=NbLig+NbLettImp ;
     end;
     ForceCommit(NbTrans) ;
     OnStoppe:=(NbLig>5000) Or (NbCpt>20) ;
     If Not FVol.Checked Then OnStoppe:=FALSE ;
     If OnStoppe Then
     begin
        GeneEnCours:=Gene ;
        AuxiEnCours:=Auxi ;
        Break ;
     end;
  end;
  If Not ShunteInitMove Then FiniMove ;
  If Not OnStoppe Then
  begin
    GeneEnCours:='' ;
    AuxiEnCours:='' ;
    FinFicRapport(F) ;
  end;
  VideListe(TLett) ;
  TLett.Free ;
  Ferme(QImp) ;
  Dispose(MvtImport) ;
END ;         *)

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 31/07/2007
Modifié le ... :   /  /    
Description .. : Permet de selectionner les écritures à lettrer
Mots clefs ... : 
*****************************************************************}
procedure TFLET.LanceLettrage(lequel : TTypeLettrage);
var
  SQL           : String ;
  Q             : TQuery ;
begin

  // Gestion des Logs :
  OuvreFicRapport(F,'',FALSE) ;

  if FVol.Checked then
  begin
     // Optimisation 
     SQL  := GetSQLSegment(lequel);
     Q    := OpenSQL(SQL,true);
     try
        // Pour chaque Tiers ou General
        while not(Q.Eof) do
        begin
           SQL := GetSQLEcriture(Q.Fields[0].AsString, lequel);
           TraiteLettrage(SQL, lequel);
           Q.Next;
        end;
     finally
        Ferme(Q);
     end;
   end
   else
   begin 
      SQL := GetSQLEcriture('', lequel);      
      TraiteLettrage(SQL, lequel);      
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 31/07/2007
Modifié le ... :   /  /    
Description .. : Regroupe les ecritures à lettrer
Mots clefs ... : 
*****************************************************************}
procedure TFLET.TraiteLettrage( SQL : string ; lequel : TTypeLettrage ) ;
var
  Q           : TQuery;
  TLett       : TList ;
  Auxi        : String;
  Gene        : String;
  Affaire     : String;
  Devise      : String;
  OkALettrer  : Boolean;    
  TotD,TotC   : Double;
  NbLettImp   : Integer;
  DevCur      : String;
  SaisieCur   : String;
  ChampMaj    : String;
  ValMaj      : String;
  CodeL       : String;
begin
  Q     := OpenSQL(SQL,True);    
  TLett := TList.Create ; 
  try
     if not ShunteInitMove then
       InitMove(RecordsCount(Q),'') ;
     while not(Q.Eof) do
     begin
        // Initialisation
        VideListe(TLett) ;
        NbLettImp   := 0 ;
        TotD        := 0 ;
        TotC        := 0 ;
        Auxi        := Q.FindField('E_AUXILIAIRE'  ).AsString ;
        Gene        := Q.FindField('E_GENERAL'     ).AsString ;
        Affaire     := Q.FindField(FNomChamp.Value ).AsString ;
        Devise      := Q.FindField('E_DEVISE'      ).AsString ;
        OkALettrer  := True ;

        while (not Q.Eof)
          and ((Q.FindField('E_AUXILIAIRE'  ).AsString = Auxi)
          and ( Q.FindField('E_GENERAL'     ).AsString = Gene)
          and ( Q.FindField(FNomChamp.Value ).AsString = Affaire)
          and ( Q.FindField('E_DEVISE'      ).AsString = Devise)) do
        begin             
           Inc(NbLettImp) ; 
           TTravail.Caption := 'Préparation ' + Gene + ' ' + Auxi + ' ' + IntToStr(NbLettImp) + ' Lignes' ;
           Application.ProcessMessages ;
           // Calcul du total Debit et Credit
           TotD := Arrondi(TotD + Q.FindField('E_DEBIT' ).AsFloat, V_PGI.OkDecV) ;
           TotC := Arrondi(TotC + Q.FindField('E_CREDIT').AsFloat, V_PGI.OkDecV) ;
           // Vérification sur la devise
           if ( DevCur <> 'PIVOT' ) and ( SaisieCur <> 'PIVOT' ) Then
           begin
              if ( DevCur = '' ) then
                 DevCur := Q.FindField('E_DEVISE').AsString ;
              if ( DevCur <> Q.FindField('E_DEVISE').AsString ) then
                 DevCur := 'PIVOT' ;
              SaisieCur := 'PIVOT' ;
           end ;

           // Sauvegarde de l'enregistrement dans la TList
           RecupMvt(Q,TLett);

           // On passe à l'enregistrement suivant
           Q.Next ;
           if not ShunteInitMove then
              MoveCur(False) ;
           
           // Si le solde est à zero on sort
           if ( lequel = lettCegid ) and
              ( Arrondi( TotD - TotC, V_PGI.OkDecV ) = 0) then Break ;
        end;



        // Test pour savoir s'il faut lettrer ou non les écritures
        case lequel of
           lettGene,lettAuxi :
           begin
              OkALettrer := ( OkALettrer  and ( NbLettImp > 1 ) and
                            (( Arrondi(TotD-TotC,V_PGI.OkDecV ) = 0 ) Or FTraitePartiel.Checked )) ;
              if CBOkDebCre.Checked then
                 OkALettrer := OkALettrer and
                               (( Arrondi( TotD, V_PGI.OkDecV ) <> 0 ) and
                                ( Arrondi( TotC, V_PGI.OkDecV ) <> 0 )) ;
           end;
           lettCegid :
           begin
              OkALettrer := OkALettrer and ( NbLettImp > 1 ) ;
              OkALettrer := OkALettrer and ((Arrondi(TotD,V_PGI.OkDecV)<>0) and
                                            (Arrondi(TotC,V_PGI.OkDecV)<>0)) ;
           end;
        end;

        if OkALettrer then
        begin                     
           VerifDevise(TLett,DevCur,SaisieCur) ;
           // Lettrage des écritures
           ChampMaj := '' ;
           ValMaj   := '' ;
           if FMaj.Checked and ( FChampMaj.Value <> '' ) and ( Trim( FValMaj.text ) <> '' ) then
           begin
              ChampMaj  := FChampMaj.Value ;
              ValMaj    := FValMaj.Text ;
           end;
           CodeL := LettrerUnPaquet(TLett,False,False,ChampMaj,ValMaj,BOptim.Checked) ;
           TTravail.Caption:='Lettrage '+Gene+' '+Auxi+' '+CodeL ;
           Application.ProcessMessages ;
           EcritFicRapport(F,Gene+' '+Auxi,CodeL,NbLettImp) ;
           DevCur:='' ;
           SaisieCur:='' ;
        end;
     end ;    
     if not ShunteInitMove then FiniMove ;   
     FinFicRapport(F) ;
  finally
     Ferme(Q);
     VideListe(TLett);
     TLett.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 31/07/2007
Modifié le ... :   /  /    
Description .. : Permet de retourner la requete à lancer sur les Tiers ou sur 
Suite ........ : les Generaux afin de segmenter la requete sur la table 
Suite ........ : Ecriture
Mots clefs ... : 
*****************************************************************}
function TFLET.GetSQLSegment( lequel : TTypeLettrage ) : string ;
var
  Champ1,Champ2     : string;
begin
  result := '';
  case lequel of
    lettAuxi,lettCegid :
    begin
       Champ1 := 'T_NATUREAUXI';
       Champ2 := 'T_AUXILIAIRE';
       result := 'SELECT T_AUXILIAIRE FROM TIERS';
    end;
    lettGene :
    begin  
       Champ1 := 'G_NATUREGENE';
       Champ2 := 'G_GENERAL';
       result := 'SELECT G_GENERAL FROM GENERAUX';
    end;
  end;

  if ( FNatCpt.ItemIndex > 0 ) then
     result := result + ' WHERE ' + Champ1 + ' = "' + FNatCpt.Value + '"';
  if ( CpteDebut.Text <> '' ) then
  begin
     if ( Pos('WHERE',result) = 0 ) then
        result := result + ' WHERE ' + Champ2 + ' >= "' + CpteDebut.Text + '"'
     else
        result := result + ' AND ' + Champ2 + ' >= "' + CpteDebut.Text + '"';
  end;
  if ( CpteFin.Text <> '' ) then
  begin
     { FQ 21344 BVE 04.09.07 }
     if ( Pos('WHERE',result) = 0 ) then
        result := result + ' WHERE ' + Champ2 + ' <= "' + CpteFin.Text + '"'
     else
        result := result + ' AND ' + Champ2 + ' <= "' + CpteFin.Text + '"';
     { END FQ 21344 }
  end;

  Result := Result + ' ORDER BY ' + Champ2;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 31/07/2007
Modifié le ... :   /  /    
Description .. : Permet de créer la requete SQL à lancer pour récuperer les 
Suite ........ : écritures à lettrer
Mots clefs ... : 
*****************************************************************}
function TFLET.GetSQLEcriture ( valeur : string ; lequel : TTypeLettrage) : string;
var
  temp : string ;  
  Lefb : tFichierBase ;
begin
  result := 'SELECT E_AUXILIAIRE, E_GENERAL, E_DEVISE, E_REFEXTERNE, E_REFLIBRE, E_AFFAIRE, ' +
            'E_LIBRETEXTE0, E_LIBRETEXTE1, E_LIBRETEXTE2, E_LIBRETEXTE3, E_LIBRETEXTE4, E_LIBRETEXTE5, ' +
            'E_LIBRETEXTE6, E_LIBRETEXTE7, E_LIBRETEXTE8, E_LIBRETEXTE9, E_REFRELEVE, E_REFLETTRAGE, E_NOMLOT, ' +
            'E_DATECOMPTABLE, E_DATEECHEANCE, E_DATEREFEXTERNE, E_LETTRAGE, E_JOURNAL, E_DEBIT, E_CREDIT, ' +
            'E_REFINTERNE, E_LIBELLE, E_DEBITDEV, E_CREDITDEV, E_NATUREPIECE, E_NUMLIGNE, E_NUMECHE, ' +
            'E_NUMEROPIECE, E_TAUXDEV ' +
            'FROM ECRITURE ';          
  if not(FVol.Checked) then
  begin
     if ( lequel = lettGene ) then
        result := result + ' LEFT JOIN GENERAUX ON E_GENERAL = G_GENERAL '
     else
        result := result + ' LEFT JOIN TIERS ON E_AUXILIAIRE = T_AUXILIAIRE ';
  end;
  result := result + ' WHERE ';
  if FVol.Checked then
  begin
     // Optimisation
     if ( lequel = lettGene ) then
        result := result + ' E_GENERAL = "' + valeur + '" '
     else
        result := result + ' E_AUXILIAIRE = "' + valeur + '" ';
  end
  else
  begin
     // Restriction sur les comptes Auxiliaires et Generaux si pas d'optimisation
     // Sinon fait dans GetSQLSegment
     if ( lequel = lettGene ) then
     begin
        result := result + ' E_AUXILIAIRE = "" ';
        if FNatCpt.ItemIndex > 0 then
           result := result + ' AND G_NATUREGENE = "' + FNatCpt.Value + '" ' ;
        if CpteDebut.Text <> '' then
           result := result + ' AND E_GENERAL >= "' + CpteDebut.Text + '" ' ;
        if CpteFin.Text <> '' then
           result := result + ' AND E_GENERAL <= "' + CpteFin.Text + '" ' ;
     end
     else
     begin
        result := result + ' E_AUXILIAIRE <> "" ';
        if FNatCpt.ItemIndex > 0 then
           result := result + ' AND T_NATUREAUXI = "' + FNatCpt.Value + '" ' ; 
        if CpteDebut.Text <> '' then
           result := result + ' AND E_AUXILIAIRE >= "' + CpteDebut.Text + '" ' ;
        if CpteFin.Text <> '' then
           result := result + ' AND E_AUXILIAIRE <= "' + CpteFin.Text + '" ' ;
     end;
  end;

  // Restriction sur les Generaux 
  if ( lequel <> lettGene ) then
  begin
     result := result + ' AND E_GENERAL <> "" ';
     if CollDebut.Text <> '' then
        result := result + ' AND E_GENERAL >= "' + CollDebut.Text + '" ' ;
     if CollFin.Text <> '' then
        result := result + ' AND E_GENERAL <= "' + CollDebut.Text + '" ' ;
  end;

  // Code de regroupement
  if ( lequel <> lettCegid ) and ( FNomChamp.Value <> '' ) then
  begin
     temp := '';
     if FCodeLettre1.Text <> '' then
        temp := ' AND ' + FNomChamp.Value + ' >= "' + FCodeLettre1.Text + '" ' ; 
     if FCodeLettre2.Text <> '' then
     begin
        if temp = '' then
           temp := ' AND ' + FNomChamp.Value + ' >= "' + FCodeLettre2.Text + '" '
        else
           temp := temp + ' AND ' + FNomChamp.Value + ' <= "' + FCodeLettre2.Text + '" ' ;
     end;
     if temp = '' then
        result := result + 'AND ' + FNomChamp.Value + ' <> "" '
     else
        result := result + temp;
  end;

  // Exercice :
  if FExercice.ITemIndex > 0 then
     result := result + ' AND E_EXERCICE="' + FExercice.Value + '" ' ;

  // Date Comptable :
  result := result + ' AND E_DATECOMPTABLE >= "' + UsDateTime(StrToDate(FDateCpta.Text)) + '" ' ;
  result := result + ' AND E_DATECOMPTABLE <= "' + UsDateTime(StrToDate(FDateCpta_.Text)) + '" ' ;
  temp := LWhereV8 ;
  if temp <> '' then
     result := result + ' AND ' + temp + ' ' ;

  // Tables Libre
  if FTypeCompte.ItemIndex=0 then
     LeFb:=fbGene
  else
     LeFb:=fbAux ;
  temp := WhereLibre(FLibre1.Text,FLibre2.Text,LeFb,TRUE) ;
  if ( temp <> '' ) then
     result := result + temp;

  // Constantes
  result := result + ' AND E_LETTRAGE = "" AND E_ETATLETTRAGE = "AL" ' +
                     ' AND E_QUALIFPIECE = "N" AND (E_ECRANOUVEAU = "N" Or E_ECRANOUVEAU = "H") ';

  // Order by
  result := result + ' ORDER BY E_AUXILIAIRE, E_GENERAL, E_DEVISE';
  if ( lequel <> lettCegid ) and ( FNomChamp.Value <> '' ) then
     result := result + ', ' + FNomChamp.Value;
  result := result + ', E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE '
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 31/07/2007
Modifié le ... :   /  /    
Description .. : Transfert les infos contenu dans un enregistrement de type 
Suite ........ : TQuey dans une TList de TL_Rappro
Mots clefs ... : 
*****************************************************************}
procedure TFLET.RecupMvt(Q : tQuery ; TLett : TList) ;
var
  L : TL_Rappro ;
begin
  if not(Q.Eof) then
  begin
    L              := TL_Rappro.Create ;
    L.General      := Q.FindField('E_GENERAL').AsString ;
    L.Auxiliaire   := Q.FindField('E_AUXILIAIRE').AsString ;
    L.DateC        := Q.FindField('E_DATECOMPTABLE').ASDateTime ;
    L.DateE        := Q.FindField('E_DATEECHEANCE').ASDateTime ;
    L.DateR        := Q.FindField('E_DATEREFEXTERNE').ASDateTime ;
    L.CodeD        := Q.FindField('E_DEVISE').AsString ;
    L.CodeL        := Q.FindField('E_LETTRAGE').AsString ;
    L.Jal          := Q.FindField('E_JOURNAL').AsString ;
    L.Debit        := Q.FindField('E_DEBIT').AsFloat ;
    L.DebitCur     := Q.FindField('E_DEBIT').AsFloat ;
    L.Credit       := Q.FindField('E_CREDIT').ASFloat ;
    L.CreditCur    := Q.FindField('E_CREDIT').ASFloat ;
    L.RefI         := Q.FindField('E_REFINTERNE').AsString ;
    L.RefE         := Q.FindField('E_REFEXTERNE').AsString ; // Correction SBO
    L.Lib          := Q.FindField('E_LIBELLE').AsString ;
    L.DebDev       := Q.FindField('E_DEBITDEV').AsFloat ;
    L.CredDev      := Q.FindField('E_CREDITDEV').ASFloat ;
    L.Nature       := Q.FindField('E_NATUREPIECE').AsString ;
    L.NumLigne     := Q.FindField('E_NUMLIGNE').AsInteger ;
    L.NumEche      := Q.FindField('E_NUMECHE').AsInteger ;
    L.Numero       := Q.FindField('E_NUMEROPIECE').AsInteger ;
    L.TauxDEV      := Q.FindField('E_TAUXDEV').ASFloat ;
    L.Decim        := V_PGI.OkDecV ;
    L.Facture      := ((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
    L.Client       := ((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
    L.Solution     := 0 ;
    L.Exo          := QuelExoDT(L.DateC) ;
    L.EditeEtatTva := False ;
    TLett.Add(L) ;
  end;
end;

(*procedure TFLET.LanceLettrageCEGIDTiers(Var GeneEnCours,AuxiEnCours : String) ;
var SQL : String ;
    QImp : TQuery ;
    TLett:TList ;
    OkALettrer : Boolean ;
    NbLettImp : integer ;
    Auxi,Gene,Devise : String ;
    MvtImport : FMvtImport ;
    TotD,TotC : Double ;
    Cpte1,Cpte2,Coll1,Coll2 : String ;
    NbTrans : Integer ;
    Date1,Date2 : TDateTime ;
    StV8 : String ;
    DevCur,SaisieCur : String ;
    OnStoppe : Boolean ;
    NbLig,NbCpt : Integer ;
    F : TextFile ;
    CodeL : String ;
    ChampMaj,ValMaj : String ;
    SQLWHERETL : String ;
    Lefb : tfichierBase ;
BEGIN
  New(MvtImport) ;
  OuvreFicRapport(F,AuxiEnCours,FALSE) ;
  OnStoppe:=FALSE ;
  NbLig:=0 ;
  NbCpt:=0 ;
  Cpte1:=CpteDebut.Text ;
  Cpte2:=CpteFin.Text ;
  NbTrans:=0 ;
  Coll1:=CollDebut.Text ;
  Coll2:=CollFin.Text ;
  Date1:=StrToDate(FDateCpta1.Text) ;
  Date2:=StrToDate(FDateCpta2.Text) ;
  SQL:='SELECT * FROM ECRITURE ' ;
  SQL:=SQL+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
  SQL:=SQL+'WHERE E_GENERAL<>"" AND '
       +'E_AUXILIAIRE<>"" AND E_LETTRAGE="" AND E_ETATLETTRAGE="AL" ' ;
  If Cpte1<>'' Then SQL:=SQL+' AND E_AUXILIAIRE>="'+Cpte1+'" ' ;
  If Cpte2<>'' Then SQL:=SQL+' AND E_AUXILIAIRE<="'+Cpte2+'" ' ;
  If Coll1<>'' Then SQL:=SQL+' AND E_GENERAL>="'+Coll1+'" ' ;
  If Coll2<>'' Then SQL:=SQL+' AND E_GENERAL<="'+Coll2+'" ' ;
  If AuxiEnCours<>'' Then SQL:=SQL+' AND E_AUXILIAIRE>="'+AuxiEnCours+'" ' ;
  If GeneEnCours<>'' Then SQL:=SQL+' AND E_GENERAL>="'+GeneEnCours+'" ' ;
  If FNatCpt.ItemIndex>0 Then SQL:=SQL+' AND T_NATUREAUXI="'+FNatCpt.Value+'" ' ;
  SQL:=SQL+' And E_DATECOMPTABLE>="'+USDATETIME(Date1)+'" ' ;
  SQL:=SQL+' And E_DATECOMPTABLE<="'+USDATETIME(Date2)+'" ' ;
  If FExercice.ITemIndex>0 Then SQL:=SQL+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
  // Correction SBO 11/08/2004 : Date testée pour ano typé H fausse !
  SQL:=SQL+' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" Or E_ECRANOUVEAU="H") ' ;
  // Fin Correction SBO
  StV8:=LWhereV8 ;
  if StV8<>'' then
     SQL:=SQL+' AND '+StV8+' ' ;

  if FTypeCOmpte.ItemIndex=0 then
     LeFb:=fbGene
  else
     LeFb:=fbAux ; 
  SQLWhereTL:=WhereLibre(FLibre1.Text,FLibre2.Text,LeFb,TRUE) ;
  If SQLWhereTL<>'' Then SQL:=SQL+SQLWHERETL ;
  SQL:=SQL+' ORDER BY E_AUXILIAIRE,E_GENERAL,E_DEVISE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ' ;
  QImp:=OpenSQL(SQL,True) ;
  if (QImp.EOF) then
  begin
     Ferme(QImp) ;
     Exit ;
  end;
  If Not ShunteInitMove Then InitMove(RecordsCount(QImp),'') ;
  TLett:=TList.Create ;
  DevCur:='' ;
  SaisieCur:='' ;
  While not QImp.Eof do
  begin
     VideListe(TLett) ;
     NbLettImp:=0 ;
     TotD:=0 ;
     TotC:=0 ;
     RecupMvtImportJacadi(Qimp,MvtImport^) ;
     Auxi:=QImp.FindField('E_AUXILIAIRE').AsString ;
     Gene:=QImp.FindField('E_GENERAL').AsString ;
     Devise:=QImp.FindField('E_DEVISE').AsString ;
     OkALettrer:=True ;
     While (not QImp.Eof) and ((QImp.FindField('E_AUXILIAIRE').AsString=Auxi)
       and (QImp.FindField('E_GENERAL').AsString=Gene)
       and (QImp.FindField('E_DEVISE').AsString=Devise)) do
     begin
        Inc(NbLettImp) ;
        TTravail.Caption:='Préparation '+Gene+' '+Auxi+' '+IntToStr(NbLettImp)+' Lignes' ;
        Application.ProcessMessages ;
        RecupMvtImportJacadi(QImp,MvtImport^) ;
        TotD:=Arrondi(TotD+MvtImport^.IE_DEBIT,V_PGI.OkDecV) ;
        TotC:=Arrondi(TotC+MvtImport^.IE_CREDIT,V_PGI.OkDecV) ;
        PrepareLettrageJacadi(TLett,MvtImport^) ;
        If (DevCur<>'PIVOT') And (SaisieCur<>'PIVOT') Then
        begin
          If DevCur='' Then DevCur:=MvtImport^.IE_DEVISE ;
          If DevCur<>MvtImport^.IE_DEVISE Then DevCur:='PIVOT' ;
          SaisieCur:='PIVOT' ;
        end;
        QImp.Next ;
        If Not ShunteInitMove Then MoveCur(False) ;
        If NbLettImp>500 Then Break ;
        If (Arrondi(TotD-TotC,V_PGI.OkDecV)=0) Then Break ;
     end;
     OkALettrer:=OkALettrer and (NbLettImp>1) ;
     OkALettrer:=OkALettrer And ((Arrondi(TotD,V_PGI.OkDecV)<>0) And (Arrondi(TotC,V_PGI.OkDecV)<>0)) ;
     if OkALettrer then
     begin
        VerifDevise(TLett,DevCur,SaisieCur) ;
        ChampMaj:='' ;
        ValMaj:='' ;
        If FMaj.Checked And (FChampMaj.Value<>'') And (Trim(FValMaj.text)<>'') Then
        begin
           ChampMaj:=FChampMaj.Value ;
           ValMaj:=FValMaj.Text ;
        end;
        CodeL:=LettrerUnPaquet(TLett,False,False,ChampMaj,ValMaj,BOptim.Checked) ;
        TTravail.Caption:='Lettrage '+Gene+' '+Auxi+' '+CodeL ;
        Application.ProcessMessages ;
        EcritFicRapport(F,Gene+' '+Auxi,CodeL,NbLettImp) ;
        DevCur:='' ;
        SaisieCur:='' ;
        Inc(NbCpt) ;
        NbLig:=NbLig+NbLettImp ;
     end;
     ForceCommit(NbTrans) ;
     OnStoppe:=(NbLig>5000) Or (NbCpt>20) ;
     If Not FVol.Checked Then OnStoppe:=FALSE ;
     If OnStoppe Then
     begin
        GeneEnCours:=Gene ;
        AuxiEnCours:=Auxi ;
        Break ;
     end;
  end;
  If Not OnStoppe Then
  begin
     GeneEnCours:='' ; AuxiEnCours:='' ;
     FinFicRapport(F) ;
  end;
  {$i-}
  CloseFile(F) ;
  {$i+}
  If Not ShunteInitMove Then FiniMove ;
  VideListe(TLett) ;
  TLett.Free ;
  Ferme(QImp) ;
  Dispose(MvtImport) ;
end;       *)

procedure TFLET.BValiderClick(Sender: TObject);
Var GeneEnCours,AuxiEnCours : String ;
begin
If Not PourRV Then if Msg.Execute(0,'','')<>mrYes then Exit ;
If PourRV Then PbRV:=FALSE ; GeneEnCours:='' ; AuxiEnCours:='' ;
EnableControls(Self,False) ;
TTravail.Caption:='Veuillez patienter...' ;
TTravail.Visible:=TRUE ; Application.ProcessMessages ;
Try
 If FMontant.Checked Then
   BEGIN
   MajChamp(FTypeCOmpte.ItemIndex=0,TRUE) ;
   If FmontantNeg.Checked Then MajChamp(FTypeCOmpte.ItemIndex=0,FALSE) ;
   END ;
 BEGINTRANS ;
 If FTypeCOmpte.ItemIndex=0 Then
   BEGIN
   Repeat
   // LanceLettrageJacadiGene(GeneEnCours,AuxiEnCours) ;
   LanceLettrage(lettGene);
   Until GeneEnCours='' ;
   END Else
   BEGIN
   Repeat
   // LanceLettrageJacadiTiers(GeneEnCours,AuxiEnCours) ;   
   LanceLettrage(lettAuxi);
   Until AuxiEnCours='' ;
   END ;
 CommitTrans ;
Except
 Rollback ;
If PourRV Then PbRV:=TRUE ;
End ;
If Not PourRV Then Msg.Execute(1,'','') ;
EnableControls(Self,True) ;
TTravail.Visible:=FALSE ;
end;

procedure TFLET.FFiltresChange(Sender: TObject);
begin
//LoadFiltre(NomFiltre,FFiltres,Pages) ;
end;

procedure TFLET.POPFPopup(Sender: TObject);
begin
//UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

procedure TFLET.FExerciceChange(Sender: TObject);
begin
If FExercice.ItemIndex>0 Then ExoToDates(FExercice.Value,FDateCpta,FDateCpta_) ;
end;

procedure TFLET.FTypeCompteClick(Sender: TObject);
Var tz : tzoomtable ;
begin
  FNatCpt.plus := '';
If FTypeCompte.ItemIndex=0 Then
  BEGIN
  tz:=tzGeneral ;
  TCollDebut.Enabled:=FALSE ; TCollFin.Enabled:=FALSE ;
  CollDebut.Enabled:=FALSE ; CollFin.Enabled:=FALSE ;
  CollDebut.Text:='' ; CollFin.Text:='' ;
// AJOUT ME Fiche 19614  FNatCpt.DataType:='ttNatGeneTIDTIC'
  FNatCpt.DataType:='ttNatGene';
  FNatCpt.plus := ' AND (CO_CODE="TID" OR CO_CODE="TIC" OR CO_CODE="DIV") ';
  END Else
  BEGIN
  tz:=tzTiers ;
  TCollDebut.Enabled:=TRUE ; TCollFin.Enabled:=TRUE ;
  CollDebut.Enabled:=TRUE ; CollFin.Enabled:=TRUE ;
  FNatCpt.DataType:='ttNatTiersCpta';
  END ;
CpteDebut.ZoomTable:=tz ; CpteFin.ZoomTable:=tz ;
end;

procedure TFLET.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then Action:=caFree ;
if Assigned(ObjetFiltre) then FreeAndNil(ObjetFiltre);
end;

procedure TFLET.FTraitePartielClick(Sender: TObject);
begin
If FTraitePartiel.Checked Then FMontant.Checked:=FALSE ;
end;

procedure TFLET.FMontantClick(Sender: TObject);
begin
If FMontant.Checked Then FTraitePartiel.Checked:=FALSE ;
FMontantNeg.Enabled:=FMontant.Checked ;
If Not FMontantNeg.Enabled Then FMontantNeg.Checked:=FALSE ;
end;

procedure TFLET.FNatCptChange(Sender: TObject);
begin
If FTypeCOmpte.ItemIndex=1 Then
  BEGIN
  Case FNatCpt.ItemIndex of
    0 : BEGIN COLLDEBUT.ZoomTable:=tzGCollectif   ; COLLFIN.ZoomTable:=tzGCollectif    ; END ;
    1 : BEGIN COLLDEBUT.ZoomTable:=tzGCollClient  ; COLLFIN.ZoomTable:=tzGCollClient   ; END ;
    2 : BEGIN COLLDEBUT.ZoomTable:=tzGCollDivers  ; COLLFIN.ZoomTable:=tzGCollDivers   ; END ;
    3 : BEGIN COLLDEBUT.ZoomTable:=tzGCollFourn   ; COLLFIN.ZoomTable:=tzGCollFourn    ; END ;
    4 : BEGIN COLLDEBUT.ZoomTable:=tzGCollSalarie ; COLLFIN.ZoomTable:=tzGCollSalarie  ; END ;
    END ;
  Case FNatCpt.ItemIndex of
    0 : BEGIN CPTEDEBUT.ZoomTable:=tzTiers      ; CPTEFIN.ZoomTable:=tzTiers      ; END ;
    1 : BEGIN CPTEDEBUT.ZoomTable:=tzTCrediteur ; CPTEFIN.ZoomTable:=tzTCrediteur ; END ;
    2 : BEGIN CPTEDEBUT.ZoomTable:=tzTDebiteur  ; CPTEFIN.ZoomTable:=tzTDebiteur  ; END ;
    3 : BEGIN CPTEDEBUT.ZoomTable:=tzTClient    ; CPTEFIN.ZoomTable:=tzTClient    ; END ;
    4 : BEGIN CPTEDEBUT.ZoomTable:=tzTDivers    ; CPTEFIN.ZoomTable:=tzTDivers    ; END ;
    5 : BEGIN CPTEDEBUT.ZoomTable:=tzTFourn     ; CPTEFIN.ZoomTable:=tzTFourn     ; END ;
    6 : BEGIN CPTEDEBUT.ZoomTable:=tzTSalarie   ; CPTEFIN.ZoomTable:=tzTSalarie   ; END ;
   end;
  END Else
  BEGIN
  Case FNatCpt.ItemIndex of
    0 : BEGIN CPTEDEBUT.ZoomTable:=TZGENERAL(*tzGTIDTIC*) ; CPTEFIN.ZoomTable:=TZGENERAL;(*tzGTIDTIC*); END ;
    //  AJOUT ME Fiche 19614
    1 : BEGIN CPTEDEBUT.ZoomTable:=TZGDIVERS      ; CPTEFIN.ZoomTable:=TZGDIVERS      ; END ;
    2 : BEGIN CPTEDEBUT.ZoomTable:=tzGTIC         ; CPTEFIN.ZoomTable:=tzGTIC         ; END ;
    3 : BEGIN CPTEDEBUT.ZoomTable:=tzGTID         ; CPTEFIN.ZoomTable:=tzGTID         ; END ;
    end;
  END ;
end;

procedure TFLET.Timer1Timer(Sender: TObject);
begin
Timer1.Enabled:=FALSE ;
BValiderClick(Nil) ;
Close ;
end;

procedure TFLET.FMAJClick(Sender: TObject);
begin
FChampMAJ.Enabled:=FMaj.Checked ; TChampMAJ.Enabled:=FMaj.Checked ;
FValMAJ.Enabled:=FMaj.Checked ; TValMAJ.Enabled:=FMaj.Checked ;
F1.Enabled:=FMaj.Checked ; F2.Enabled:=FMaj.Checked ;
If Not FMaj.Checked Then BEGIN FValMAJ.Text:='' ; FChampMaj.ItemIndex:=-1 ; END ;
end;

Procedure TFLET.UpdateLeLettrageCegid  ;
var SQL : String ;
    QImp : TQuery ;
    Cpte1,Cpte2,Coll1,Coll2 : String ;
    NbTrans : Integer ;
    Date1,Date2 : TDateTime ;
    StV8 : String ;
    C,CD : Double ;
    OkMaj : Boolean ;
BEGIN
Exit ;
Cpte1:=CpteDebut.Text ; Cpte2:=CpteFin.Text ; NbTrans:=0 ;
Coll1:=CollDebut.Text ; Coll2:=CollFin.Text ;
Date1:=StrToDate(FDateCpta.Text) ; Date2:=StrToDate(FDateCpta_.Text) ;
SQL:='SELECT E_DEBIT,E_CREDIT,E_DEBITDEV,E_CREDITDEV,E_ETAT,E_DATEPAQUETMIN,E_DATEPAQUETMAX,E_DATECOMPTABLE, '+
     'E_COUVERTURE,E_COUVERTUREDEV,E_LETTRAGE,E_ETATLETTRAGE,E_LETTRAGEDEV, E_TRESOSYNCHRO FROM ECRITURE ';{JP 26/04/04 : pour l'échéancier de la Tréso}
SQL:=SQL+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
SQL:=SQL+'WHERE E_GENERAL<>"" AND '
     +'E_AUXILIAIRE<>"" AND '+FChampMAJ.Value+'="'+FValMaj.TExt+'" AND E_ETATLETTRAGE="PL" ' ;
If Cpte1<>'' Then SQL:=SQL+' AND E_AUXILIAIRE>="'+Cpte1+'" ' ;
If Cpte2<>'' Then SQL:=SQL+' AND E_AUXILIAIRE<="'+Cpte2+'" ' ;
If Coll1<>'' Then SQL:=SQL+' AND E_GENERAL>="'+Coll1+'" ' ;
If Coll2<>'' Then SQL:=SQL+' AND E_GENERAL<="'+Coll2+'" ' ;
If FNatCpt.ItemIndex>0 Then SQL:=SQL+' AND T_NATUREAUXI="'+FNatCpt.Value+'" ' ;
SQL:=SQL+' And E_DATECOMPTABLE>="'+USDATETIME(Date1)+'" ' ;
SQL:=SQL+' And E_DATECOMPTABLE<="'+USDATETIME(Date2)+'" ' ;
If FExercice.ITemIndex>0 Then SQL:=SQL+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
// Correction SBO 11/08/2004 : Date testée pour ano typé H fausse !
SQL:=SQL+' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" Or E_ECRANOUVEAU="H") ' ;
// Fin Correction SBO
StV8:=LWhereV8 ; if StV8<>'' then SQL:=SQL+' AND '+StV8+' ' ;
SQL:=SQL+' ORDER BY E_AUXILIAIRE,E_GENERAL,E_DEVISE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ' ;
QImp:=OpenSQL(SQL,FALSE) ;
if (QImp.EOF) then BEGIN Ferme(QImp) ; Exit ; END ;
If Not ShunteInitMove Then InitMove(RecordsCount(QImp),'') ;
While not QImp.Eof do
  BEGIN
  C:=QImp.FindField('E_COUVERTURE').AsFloat ; CD:=QImp.FindField('E_COUVERTUREDEV').AsFloat ;
  OkMaj:=(Arrondi(C,V_PGI.OkDecV)=0) Or (Arrondi(CD,4)=0);
  If OkMaj Then
    BEGIN
    QImp.Edit ;
    QImp.FindField('E_COUVERTURE').AsFloat:=0 ;
    QImp.FindField('E_COUVERTUREDEV').AsFloat:=0 ;
    QImp.FindField('E_ETAT').AsString:='0000000000' ;
    QImp.FindField('E_ETATLETTRAGE').AsString:='AL' ;
    QImp.FindField('E_LETTRAGE').AsString:='' ;
    QImp.FindField('E_LETTRAGEDEV').AsString:='-' ;
    QImp.FindField('E_DATEPAQUETMIN').AsString:=QImp.FindField('E_DATECOMPTABLE').AsString ;
    QImp.FindField('E_DATEPAQUETMAX').AsString:=QImp.FindField('E_DATECOMPTABLE').AsString ;
    QImp.FindField('E_TRESOSYNCHRO').AsString := 'LET'; {JP 26/04/04 : pour l'échéancier de la Tréso}
    QImp.Post ;
    END ;
  QImp.Next ; If Not ShunteInitMove Then MoveCur(False) ;
  ForceCommit(NbTrans) ;
  END ;
If Not ShunteInitMove Then FiniMove ;
Ferme(QImp) ;
END ;


procedure TFLET.BValiderCEGIDClick(Sender: TObject);
Var GeneEnCours,AuxiEnCours : String ;
    Pb : Boolean ;
begin
If Not PourRV Then if Msg.Execute(0,'','')<>mrYes then Exit ;
If PourRV Then PbRV:=FALSE ;
GeneEnCours:='' ; AuxiEnCours:='' ; Pb:=FALSE ;
EnableControls(Self,False) ;
TTravail.Caption:='Veuillez patienter...' ;
TTravail.Visible:=TRUE ; Application.ProcessMessages ;
Try
 BEGINTRANS ;
   Repeat
   // LanceLettrageCEGIDTiers(GeneEnCours,AuxiEnCours) ;
   LanceLettrage(lettAuxi);
   Until AuxiEnCours='' ;
 CommitTrans ;
Except
 Rollback ;
 If PourRV Then PbRV:=TRUE ;
 Pb:=TRUE ;
End ;
If Not Pb Then
  BEGIN
  Try
   BEGINTRANS ;
   GeneEnCours:='' ; AuxiEnCours:='' ;
   Repeat
   UpdateLeLettrageCegid ;
   Until GeneEnCours='' ;
   CommitTrans ;
  Except
   Rollback ;
  End ;
  END ;
If Not PourRV Then Msg.Execute(1,'','') ;
EnableControls(Self,True) ;
TTravail.Visible:=FALSE ;
end;

procedure TFLET.FLibre1DblClick(Sender: TObject);
Var StTri, StCod1, StCod2, Pref : String ;
    LeFb : TFichierBase ;
begin
StCod1:=FLibre1.text ; StCod2:=FLibre2.text ;
if FTypeCOmpte.ItemIndex=0 then BEGIN LeFb:=fbGene ; Pref:='G0' ; END Else BEGIN LeFb:=fbAux ; Pref:='T0' ; END ;
StTri:='' ;
ChoixTableLibreSur(LeFb,StTri, StCod1, StCod2) (*ChoixTableLibre(LeFb,StTri, StCod1, StCod2)*) ;
FLibre1.text:=StCod1 ; FLibre2.text:=StCod2 ;
End ;

procedure TFLET.BDeletClick(Sender: TObject);
var SQL : String ;
    QImp : TQuery ;
    Cpte1,Cpte2,Coll1,Coll2 : String ;
    Date1,Date2 : TDateTime ;
    StV8 : String ;
    OkMaj : Boolean ;
BEGIN
if Msg.Execute(2,'','')<>mrYes then Exit ;
Cpte1:=CpteDebut.Text ; Cpte2:=CpteFin.Text ;
Coll1:=CollDebut.Text ; Coll2:=CollFin.Text ;
Date1:=StrToDate(FDateCpta.Text) ; Date2:=StrToDate(FDateCpta_.Text) ;
SQL:='SELECT E_DEBIT,E_CREDIT,E_DEBITDEV,E_CREDITDEV,E_ETAT,E_DATEPAQUETMIN,E_DATEPAQUETMAX,E_DATECOMPTABLE, '+
     'E_COUVERTURE,E_COUVERTUREDEV,E_LETTRAGE,E_ETATLETTRAGE,E_LETTRAGEDEV,'+FChampMAJ.Value+', E_TRESOSYNCHRO FROM ECRITURE ';{JP 26/04/04 : pour l'échéancier de la Tréso}
SQL:=SQL+'WHERE E_GENERAL<>"" AND '
     +'E_AUXILIAIRE<>"" AND '+FChampMAJ.Value+'="'+FValMaj.TExt+'" AND E_ETATLETTRAGE="PL" ' ;
If Cpte1<>'' Then SQL:=SQL+' AND E_AUXILIAIRE>="'+Cpte1+'" ' ;
If Cpte2<>'' Then SQL:=SQL+' AND E_AUXILIAIRE<="'+Cpte2+'" ' ;
If Coll1<>'' Then SQL:=SQL+' AND E_GENERAL>="'+Coll1+'" ' ;
If Coll2<>'' Then SQL:=SQL+' AND E_GENERAL<="'+Coll2+'" ' ;
//If FNatCpt.ItemIndex>0 Then SQL:=SQL+' AND T_NATUREAUXI="'+FNatCpt.Value+'" ' ;
SQL:=SQL+' And E_DATECOMPTABLE>="'+USDATETIME(Date1)+'" ' ;
SQL:=SQL+' And E_DATECOMPTABLE<="'+USDATETIME(Date2)+'" ' ;
If FExercice.ITemIndex>0 Then SQL:=SQL+' AND E_EXERCICE="'+FExercice.Value+'" ' ;
// Correction SBO 11/08/2004 : Date testée pour ano typé H fausse !
SQL:=SQL+' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" Or E_ECRANOUVEAU="H") ' ;
// Fin Correction SBO
StV8:=LWhereV8 ; if StV8<>'' then SQL:=SQL+' AND '+StV8+' ' ;
SQL:=SQL+' ORDER BY E_AUXILIAIRE,E_GENERAL,E_DEVISE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ' ;
QImp:=OpenSQL(SQL,FALSE) ;
if (QImp.EOF) then BEGIN Ferme(QImp) ; Exit ; END ;
If Not ShunteInitMove Then InitMove(RecordsCount(QImp),'') ;
While not QImp.Eof do
  BEGIN
  (*
  C:=QImp.FindField('E_COUVERTURE').AsFloat ; CD:=QImp.FindField('E_COUVERTUREDEV').AsFloat ; CE:=QImp.FindField('E_COUVERTUREEURO').AsFloat ;
  OkMaj:=(Arrondi(C,V_PGI.OkDecV)=0) Or (Arrondi(CD,4)=0) Or (Arrondi(CE,V_PGI.OkDecE)=0) ;
  *)
  OkMaj:=TRUE ;
  If OkMaj Then
    BEGIN
    QImp.Edit ;
    QImp.FindField('E_COUVERTURE').AsFloat:=0 ;
    QImp.FindField('E_COUVERTUREDEV').AsFloat:=0 ;
    QImp.FindField('E_ETAT').AsString:='0000000000' ;
    QImp.FindField('E_ETATLETTRAGE').AsString:='AL' ;
    QImp.FindField('E_LETTRAGE').AsString:='' ;
    QImp.FindField('E_LETTRAGEDEV').AsString:='-' ;
    QImp.FindField(FChampMAJ.Value).AsString:='' ;
    QImp.FindField('E_DATEPAQUETMIN').AsString:=QImp.FindField('E_DATECOMPTABLE').AsString ;
    QImp.FindField('E_DATEPAQUETMAX').AsString:=QImp.FindField('E_DATECOMPTABLE').AsString ;
    QImp.FindField('E_TRESOSYNCHRO').AsString := 'LET'; {JP 26/04/04 : pour l'échéancier de la Tréso}
    QImp.Post ;
    END ;
  QImp.Next ; If Not ShunteInitMove Then MoveCur(False) ;
  END ;
If Not ShunteInitMove Then FiniMove ;
Ferme(QImp) ;
END ;
         
{ FQ 21262 BVE 21.08.07 }
procedure TFLET.BFermeClick(Sender: TObject);
begin
  if (IsInside(Self)) then CloseInsidePanel(self);
  close;
end;

procedure TFLET.FDateCptaKeyPress(Sender: TObject; var Key: Char);
begin
  ParamDate(self,sender,key);
end;

{ END FQ 21262 }

end.
