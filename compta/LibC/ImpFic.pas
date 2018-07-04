unit ImpFic;

interface
uses
  Windows, SysUtils, Classes, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls,
{$IFNDEF EAGLCLIENT}
  PrintDBG,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
  UTOB,
{$ENDIF}
  Ent1, HSysMenu,
  hmsgbox, HEnt1, HTB97, VisuEnr, RappType,ImpFicU,TImpFic,
{$IFNDEF PROGEXT}
  RapSuppr,
{$ENDIF}

  Menus ;

Procedure VisuLignesErreurs (StFichier,StErr : String ; FormatFic : integer ; Bourre : Boolean ; ignorel1 : boolean=False) ;

Procedure ChargeJalSoc(Quoi : Integer ; Var JalSoc : TTabCptLu) ;
Function  TrouveJalSoc(Cpt : String ; Var TabCpt : TTabCptLu) : TCptLu ;
Function  NbCptLu(Var TabCpt : TTabCptLu) : Integer ;
{$IFNDEF EAGLCLIENT}
Procedure FicToFicCorresp(Quoi,FormatFic : Integer ; StFichier,StNewFichier1 : String) ;
{$ENDIF}
Procedure AlimEnr(St : String ; Var Enr : TStLue ; FormatFic : Integer) ;


type
  TFCImpFic = class(TForm)
    Outils: TPanel;
    HMTrad: THSystemMenu;
    G: THGrid;
    NbLignes: TLabel;
    NbLignesFausses: TLabel;
    HFormatSAARI: THMsgBox;
    HFormatHALLEY: THMsgBox;
    Patience: TLabel;
    FindMvt: TFindDialog;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    Uutils97: TToolbar97;
    BImprimer: TToolbarButton97;
    BRecherche: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    Valide97: TToolbar97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    GP: THGrid;
    Sauve: TSaveDialog;
    HMess: THMsgBox;
    NbPieces: TLabel;
    NbPiecesFausses: TLabel;
    NbPiecesIncorrectes: TLabel;
    HFormatCGN: THMsgBox;
    BMenuZoom: TToolbarButton97;
    POPZ: TPopupMenu;
    TranfertFic: TBitBtn;
    VerifCpt: TBitBtn;
    VisuCpt: TBitBtn;
    GPL: THGrid;
    Visulignespiece: TBitBtn;
    VisuLignes: TBitBtn;
    VisuTotaux: TBitBtn;
    VisuLignesFausses: TBitBtn;
    VisuBlancs: TBitBtn;
    VoirBlanc: TCheckBox;
    CalcNum: TBitBtn;
    Corresp: TBitBtn;
    CalcTC: TBitBtn;
    CalcAttent: TBitBtn;
    HFormatCGE: THMsgBox;
    GAux: THGrid;
    VisuAux: TBitBtn;
    GGen: THGrid;
    VisuGen: TBitBtn;
    GSect: THGrid;
    VisuGerer1: TBitBtn;
    WhatError: TLabel;
    VisuGener2: TBitBtn;
    VisuGener4: TBitBtn;
    VisuGener5: TBitBtn;
    VisuGener6: TBitBtn;
    VisuExerc: TBitBtn;
    G7: THGrid;
    VisuSect: TBitBtn;
    G8: THGrid;
    G9: THGrid;
    G10: THGrid;
    G11: THGrid;
    G12: THGrid;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BImprimerClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure FindMvtFind(Sender: TObject);
    procedure VerifCptClick(Sender: TObject);
    Procedure RecopieFic(StNumP : String) ;
    procedure FormCreate(Sender: TObject);
    procedure GDblClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure TranfertFicClick(Sender: TObject);
    procedure VisuCptClick(Sender: TObject);
    procedure VisuTotauxClick(Sender: TObject);
    procedure VisuLignesClick(Sender: TObject);
    procedure VisulignespieceClick(Sender: TObject);
    procedure VisuLignesFaussesClick(Sender: TObject);
    procedure VisuBlancsClick(Sender: TObject);
    procedure CalcNumClick(Sender: TObject);
    procedure CorrespClick(Sender: TObject);
    procedure CalcTCClick(Sender: TObject);
    procedure CalcAttentClick(Sender: TObject);
    procedure VisuAuxClick(Sender: TObject);
    procedure VisuGenClick(Sender: TObject);
    procedure VisuGerer1Click(Sender: TObject);
    procedure VisuGerer7Click(Sender: TObject);
    procedure GRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure VisuGener2Click(Sender: TObject);
    procedure VisuGener4Click(Sender: TObject);
    procedure VisuGener5Click(Sender: TObject);
    procedure VisuGener6Click(Sender: TObject);
    procedure VisuExercClick(Sender: TObject);
  private
    Function ChargeGrid : Integer ;
    procedure AlimGrid(Enr : TStLue ; Lig,FormatFic : Integer) ;
    Function  AlimGridCpt(What : Integer ; Var St : String ; Lig : Integer) : Boolean ;
    Function  VerifEnr(Var Enr : TStLue) : Integer ;
    procedure InitGrid ;
    Function  VerifCompte(What : Integer ; Var TabCpt : TTabCptLu) : Boolean ;
    Function GereTotalpiece(WhatErreur : Integer ; OldEnr,Enr : TStLue ; Var TotPieceD,TotPieceC : Double ; Var NumP,NumPE,NumPS : Integer) : String ;
    Procedure AlimCptLu(Enr : TStLue) ;
    Function  QuelMessSup(What : Integer ; CptLu : TCptLu) : String ;
    Procedure RecupEnr(Var Enr : TStLue ; VoirLesBlancs : Boolean ; Lig : Integer) ;
    Procedure QuelGrid(Quoi : Byte) ;
    Procedure MasqueLigne(Quoi : Integer) ;
    procedure VoirOuNonLesBlancs;
  // Function  FindFirstLig(StNumP : String ; Avance : Boolean) : Integer ;
 //   procedure AlimGPL(Lig : Integer) ;
    procedure VisibleGrid (Quoi : Byte);
  public
    Ignorel1  : Boolean ;
    StFichier : String ;
    FormatFic : Integer ;
    FindFirst : Boolean ;
    RechCompteLance : Boolean ;
    StNumP    : String ;
    GenLu     : TTabCptLu ;
    AuxLu     : TTabCptLu ;
    SecLu     : TTabCptLu ;
    ListeCptFaux : TList ;
    ModifFaite : Boolean ;
    Bourre     : Boolean ;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  HStatus;


Procedure VisuLignesErreurs (StFichier,StErr : String ; FormatFic : integer ; Bourre : Boolean ; ignorel1 : boolean=False) ;
Var X : TFCImpFic ;
BEGIN
X:=TFCImpFic.Create(Application) ;
 Try
  X.Ignorel1:=Ignorel1 ;
  X.StFichier:=StFichier ;
  X.FormatFic:=FormatFic ;
  X.Bourre:=Bourre ;
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
END ;

Function MessErreurCompte(What : Integer ; CptLu : TCptLu) : String ;
Var Racine : String ;

BEGIN
Result:='' ;
Case What Of
  0 : Result:='Le compte général n''existe pas.' ;
  1 : Result:='Le compte auxiliaire n''existe pas.' ;
  2 : Result:='La section analytique n''existe pas.' ;
  END ;

Case What Of
  0 : BEGIN
      Racine:=Copy(CptLu.Cpt,1,1) ;
      If Racine='6' Then  Result:=Result+' Compte de charge' Else
         If Racine='7' Then  Result:=Result+' Compte de produit' ;
      If CptLu.Ventilable Then Result:=Result+' Ventilable' ;
      END ;
  1 : BEGIN
      Racine:=Copy(CptLu.Collectif,1,2) ;
      If Racine='40' Then Result:=Result+' Fournisseur' Else
         If Racine='41' Then Result:=Result+' Client' ;
      If CptLu.Collectif<>'' Then Result:=Result+' - '+CptLu.Collectif ;
      END ;
  END ;
Result:=TraduireMemoire(Result) ;
END ;



Function VoirBlancs(St : String ; Voir : Boolean) : String ;
Var i : Integer ;
BEGIN
If Voir Then
   BEGIN
   For i:=1 To Length(St) Do If St[i]=' ' Then St[i]:='_' ;
   END Else
   BEGIN
   For i:=1 To Length(St) Do If St[i]='_' Then St[i]:=' ' ;
   END ;
Result:=st ;
END ;

Function OkEntier(Var St : String ; AvecTrim : Boolean) : Boolean ;
Var StTest : String ;
BEGIN
Result:=TRUE ; If AvecTrim Then StTest:=Trim(St) Else StTest:=St ;
If (StTest<>'') And (Not VerifEntier(StTest)) Then BEGIN St:='!'+St+'!' ; RESULT:=FALSE ; END ;
END ;

Function OkVide(Var St : String) : Boolean ;
BEGIN
Result:=TRUE ;
If Trim(St)='' Then BEGIN St:='!'+St+'!' ; Result:=FALSE ; END ;
END ;

Procedure AlimEnr(St : String ; Var Enr : TStLue ; FormatFic : Integer) ;
BEGIN
Fillchar(Enr,SizeOf(Enr),#0) ;
Case FormatFic Of
  0 : BEGIN { SAARI COEUR DE GAMME ET NEGOCE V2 }
      Enr.Jal:=Copy(St,1,3) ; Enr.Date:=Copy(St,4,6) ; Enr.TP:=Copy(St,10,2) ;
      Enr.General:=Copy(St,12,13) ; Enr.TC:=Copy(St,25,1) ; Enr.AuxSect:=Copy(St,26,13) ;
      Enr.Reference:=Copy(St,39,13) ; Enr.Libelle:=Copy(St,52,25) ; Enr.MP:=Copy(St,77,1) ;
      Enr.Echeance:=Copy(St,78,6) ; Enr.S:=Copy(St,84,1) ; Enr.Montant:=Copy(St,85,20) ;
      Enr.TE:=Copy(St,105,1) ; Enr.NumP:=Copy(St,106,7) ;
      If Sn2Orli Then
         BEGIN
         Enr.MontantD:=Copy(St,113,20) ; Enr.Dev:=Copy(St,133,3) ;
         Enr.TauxE:=Copy(St,136,10) ; Enr.MPaie:=Copy(St,146,3) ;
         Enr.Rib:=Copy(St,149,50) ;
         END ;
      END ;
  1 : BEGIN { HALLEY LIGHT }
      Enr.Jal:=Copy(St,1,3) ; Enr.Date:=Copy(St,4,8) ; Enr.TP:=Copy(St,12,2) ;
      Enr.General:=Copy(St,14,17) ; Enr.TC:=Copy(St,31,1) ; Enr.AuxSect:=Copy(St,32,17) ;
      Enr.Reference:=Copy(St,49,13) ; Enr.Libelle:=Copy(St,62,25) ; Enr.MP:=Copy(St,87,1) ;
      Enr.Echeance:=Copy(St,88,8) ; Enr.S:=Copy(St,96,1) ; Enr.Montant:=Copy(St,97,20) ;
      Enr.TE:=Copy(St,117,1) ; Enr.NumP:=Copy(St,118,7) ;
      END ;
  2 : BEGIN { HALLEY ETENDU }
      Enr.Jal:=Copy(St,1,3) ; Enr.Date:=Copy(St,4,8) ; Enr.TP:=Copy(St,12,2) ;
      Enr.General:=Copy(St,14,17) ; Enr.TC:=Copy(St,31,1) ; Enr.AuxSect:=Copy(St,32,17) ;
      Enr.Reference:=Copy(St,49,13) ; Enr.Libelle:=Copy(St,62,25) ; Enr.MP:=Copy(St,87,1) ;
      Enr.Echeance:=Copy(St,88,8) ; Enr.S:=Copy(St,96,1) ; Enr.Montant:=Copy(St,97,20) ;
      Enr.TE:=Copy(St,117,1) ; Enr.NumP:=Copy(St,118,7) ; Enr.NumEche:=Copy(St,125,2) ;
      Enr.Valid:=Copy(St,127,1) ; Enr.RefExt:=Copy(St,128,35) ; Enr.DateRefExt:=Copy(St,163,8) ;
      Enr.DateCreat:=Copy(St,171,8) ; Enr.DateModif:=Copy(St,179,8) ; Enr.Soc:=Copy(St,187,3) ;
      Enr.Etab:=Copy(St,190,3) ; Enr.Affaire:=Copy(St,193,17) ; Enr.DebE:=Copy(St,210,20) ;
      Enr.CreE:=Copy(St,230,20) ; Enr.TauxE:=Copy(St,250,20) ; Enr.Dev:=Copy(St,270,3) ;
      Enr.DebD:=Copy(St,273,20) ; Enr.CreD:=Copy(St,293,20) ; Enr.TauxD:=Copy(St,313,20) ;
      Enr.DateTauxD:=Copy(St,333,8) ; Enr.Quot:=Copy(St,341,20) ; Enr.EcrAN:=Copy(St,361,3) ;
      Enr.Qte1:=Copy(St,364,20) ; Enr.Qte2:=Copy(St,384,20) ; Enr.Qual1:=Copy(St,404,3) ;
      Enr.Qual2:=Copy(St,407,3) ; Enr.RefLibre:=Copy(St,410,35) ; Enr.TvaEnc:=Copy(St,445,1) ;
      Enr.Regime:=Copy(St,446,3) ; Enr.Tva:=Copy(St,449,3) ; Enr.TPF:=Copy(St,452,3) ;
      Enr.CtrGen:=Copy(St,455,17) ; Enr.CtrAux:=Copy(St,472,17) ; Enr.Couv:=Copy(St,489,20) ;
      Enr.Let:=Copy(St,509,5) ; Enr.LetD:=Copy(St,514,1) ; Enr.DateMax:=Copy(St,515,8) ;
      Enr.DateMin:=Copy(St,523,8) ; Enr.RefP:=Copy(St,531,17) ; Enr.DateP:=Copy(St,548,8) ;
      Enr.LPLCR:=Copy(St,556,4) ; Enr.DateRel:=Copy(St,560,8) ; Enr.Controle:=Copy(St,568,1) ;
      Enr.TotEnc:=Copy(St,569,20) ; Enr.RelEnc:=Copy(St,589,20) ; Enr.DateV:=Copy(St,609,8) ;
      Enr.RIB:=Copy(St,617,35) ; Enr.RefRel:=Copy(St,652,35) ; Enr.CouvD:=Copy(St,687,20) ;
      Enr.EtatL:=Copy(St,707,3) ; Enr.NumPInt:=Copy(St,710,35) ; Enr.Encais:=Copy(St,745,3) ;
      Enr.TAN:=Copy(St,748,3) ; Enr.Eche:=Copy(St,751,1) ; Enr.Ana:=Copy(St,752,1) ;
      Enr.MPaie:=Copy(St,753,3) ; Enr.NatP:=Copy(St,756,3) ;
      END ;
  3,4 : BEGIN { CEGID NORMAL }
      Enr.Jal:=Copy(St,1,3) ; Enr.Date:=Copy(St,4,8) ; Enr.TP:=Copy(St,12,2) ;
      Enr.General:=Copy(St,14,17) ; Enr.TC:=Copy(St,31,1) ; Enr.AuxSect:=Copy(St,32,17) ;
      Enr.Reference:=Copy(St,49,35) ; Enr.Libelle:=Copy(St,84,35) ; Enr.MP:=Copy(St,119,3) ;
      Enr.Echeance:=Copy(St,122,8) ; Enr.S:=Copy(St,130,1) ; Enr.Montant:=Copy(St,131,20) ;
      Enr.TE:=Copy(St,151,1) ; Enr.NumP:=Copy(St,152,8) ; Enr.Dev:=Copy(St,160,3) ;
      Enr.TauxE:=Copy(St,163,10) ; Enr.CodeMontant:=Copy(St,173,3) ; Enr.MontantD:=Copy(St,176,20) ;
      Enr.MontantE:=Copy(St,196,20) ; Enr.Etab:=Copy(St,216,3) ;Enr.Axe:=Copy(St,219,2) ;Enr.NumEche:=Copy(St,221,2) ;
      If FormatFic=4 Then
        BEGIN
        Enr.RefExt:=Copy(St,223,35) ; Enr.DateRefExt:=Copy(St,258,8) ; Enr.DateCreat:=Copy(St,266,8) ;
        Enr.Soc:=Copy(St,274,3)     ; Enr.Affaire:=Copy(St,277,17)   ; Enr.DateTauxD:=Copy(St,294,8) ;
        Enr.EcrAN:=Copy(St,302,3)   ; Enr.Qte1:=Copy(St,305,20)      ; Enr.Qte2:=Copy(St,325,20) ;
        Enr.Qual1:=Copy(St,345,3)   ; Enr.Qual2:=Copy(St,348,3)      ; Enr.RefLibre:=Copy(St,351,35) ;
        Enr.TvaEnc:=Copy(St,386,1)  ; Enr.Regime:=Copy(St,387,3)     ; Enr.Tva:=Copy(St,390,3) ;
        Enr.TPF:=Copy(St,393,3)     ; Enr.CtrGen:=Copy(St,396,17)    ; Enr.CtrAux:=Copy(St,413,17) ;
        Enr.RefP:=Copy(St,430,17)   ; Enr.DateP:=Copy(St,447,8)      ; Enr.DateRel:=Copy(St,455,8) ;
        Enr.DateV:=Copy(St,463,8)   ; Enr.RIB:=Copy(St,471,35)       ; Enr.RefRel:=Copy(St,506,10);
        Enr.Immo:=Copy(St,516,17)   ; Enr.LT0:=Copy(St,533,30)       ; Enr.LT1:=Copy(St,563,30) ;
        Enr.LT2:=Copy(St,593,30)    ; Enr.LT3:=Copy(St,623,30)       ; Enr.LT4:=Copy(St,653,30) ;
        Enr.LT5:=Copy(St,683,30)    ; Enr.LT6:=Copy(St,713,30)       ; Enr.LT7:=Copy(St,743,30) ;
        Enr.LT8:=Copy(St,773,30)    ; Enr.LT9:=Copy(St,803,30)       ; Enr.TA0:=Copy(St,833,3) ;
        Enr.TA1:=Copy(St,836,3)     ; Enr.TA2:=Copy(St,839,3)        ; Enr.TA3:=Copy(St,842,3) ;
        Enr.LM0:=Copy(St,845,20)    ; Enr.LM1:=Copy(St,865,20)       ; Enr.LM2:=Copy(St,885,20) ;
        Enr.LM3:=Copy(St,905,20)    ; Enr.LD:=Copy(St,925,8)         ; Enr.LB0:=Copy(St,933,1) ;
        Enr.LB1:=Copy(St,934,1)     ; Enr.Conso:=Copy(St,935,3)      ; Enr.Couv:=Copy(St,938,20) ;
        Enr.CouvD:=Copy(St,958,20)  ; Enr.CouvE:=Copy(St,978,20)     ; Enr.DateMax:=Copy(St,998,8) ;
        Enr.DateMin:=Copy(St,1006,8); Enr.Let:=Copy(St,1014,5)       ; Enr.LetD:=Copy(St,1019,1) ;
        Enr.LetE:=Copy(St,1020,1)   ; Enr.EtatL:=Copy(St,1021,3) ;
        END ;
      END ;
  END ;
END ;

Function TFCImpFic.VerifEnr(Var Enr : TStLue) : Integer ;
Var UneErreur : Integer ;
    StTest : String ;
{
  1  : N° pièce incorrect
  2  : Date pièce non renseigné
  3  : Date pièce vide
  4  : Jal non renseigné
  5  : Compte général non renseigné
  6  : Type écriture non renseigné
  7  : Montant incorrect
  8  : Sens incorrect
  9  : Qualif incorrect
  10 : Type compte incorrect
}
BEGIN
UneErreur:=0 ;
If Not OkEntier(Enr.NumP,TRUE) Then UneErreur:=1 ;
If OkEntier(Enr.Date,FALSE) Then
   BEGIN
   If Not OkVide(Enr.Date) Then UneErreur:=2 ;
   END Else UneErreur:=3 ;
If Not OkVide(Enr.Jal) Then UneErreur:=4 ;
If Not OkVide(Enr.General) Then UneErreur:=5 ;
If Not OkVide(Enr.TE) Then UneErreur:=6 ;
StTest:=Trim(Enr.Montant) ;
If (StTest<>'') And (Not VerifMontant(StTest)) Then BEGIN Enr.Montant:='!'+Enr.Montant+'!' ; UneErreur:=7 ; END ;
If (Enr.S[1] In ['D','C'])=FALSE Then BEGIN Enr.S:='!'+Enr.S+'!' ; UneErreur:=8 ; END ;
If (Enr.TE[1] In ['N','S','A'])=FALSE Then BEGIN Enr.TE:='!'+Enr.TE+'!' ; UneErreur:=9 ; END ;
If (Enr.TC[1] In ['X','A',' ','E','H','O','L'])=FALSE Then BEGIN Enr.TC:='!'+Enr.TC+'!' ; UneErreur:=10 ; END ;
Result:=UneErreur ;
END ;

procedure TFCImpFic.AlimGrid(Enr : TStLue ; Lig,FormatFic : Integer) ;
BEGIN
G.Cells[0,Lig]:=Enr.NumLg ;

G.Cells[1,Lig]:=VoirBlancs(Enr.Jal,VoirBlanc.Checked) ;
G.Cells[2,Lig]:=VoirBlancs(Enr.Date,VoirBlanc.Checked) ;
G.Cells[3,Lig]:=VoirBlancs(Enr.TP,VoirBlanc.Checked) ;
G.Cells[4,Lig]:=VoirBlancs(Enr.General,VoirBlanc.Checked) ;
G.Cells[5,Lig]:=VoirBlancs(Enr.TC,VoirBlanc.Checked) ;
G.Cells[6,Lig]:=VoirBlancs(Enr.AuxSect,VoirBlanc.Checked) ;
G.Cells[7,Lig]:=VoirBlancs(Enr.Reference,VoirBlanc.Checked) ;
G.Cells[8,Lig]:=VoirBlancs(Enr.Libelle,VoirBlanc.Checked) ;
G.Cells[9,Lig]:=VoirBlancs(Enr.MP,VoirBlanc.Checked) ;
G.Cells[10,Lig]:=VoirBlancs(Enr.Echeance,VoirBlanc.Checked) ;
G.Cells[11,Lig]:=VoirBlancs(Enr.S,VoirBlanc.Checked) ;
G.Cells[12,Lig]:=VoirBlancs(Enr.Montant,VoirBlanc.Checked) ;
G.Cells[13,Lig]:=VoirBlancs(Enr.TE,VoirBlanc.Checked) ;
G.Cells[14,Lig]:=VoirBlancs(Enr.NumP,VoirBlanc.Checked) ;
If Sn2Orli Then
   BEGIN
   G.Cells[15,Lig]:=VoirBlancs(Enr.MontantD,VoirBlanc.Checked) ;
   G.Cells[16,Lig]:=VoirBlancs(Enr.Dev,VoirBlanc.Checked) ;
   G.Cells[17,Lig]:=VoirBlancs(Enr.TauxE,VoirBlanc.Checked) ;
   G.Cells[18,Lig]:=VoirBlancs(Enr.MPaie,VoirBlanc.Checked) ;
   G.Cells[19,Lig]:=VoirBlancs(Enr.Rib,VoirBlanc.Checked) ;
   END Else
   Case FormatFic Of
     0,1,2 : BEGIN
             G.Cells[15,Lig]:=VoirBlancs(Enr.NumEche,VoirBlanc.Checked) ;
             G.Cells[16,Lig]:=VoirBlancs(Enr.Valid,VoirBlanc.Checked) ;
             G.Cells[17,Lig]:=VoirBlancs(Enr.RefExt,VoirBlanc.Checked) ;
             G.Cells[18,Lig]:=VoirBlancs(Enr.DateRefExt,VoirBlanc.Checked) ;
             G.Cells[19,Lig]:=VoirBlancs(Enr.DateCreat,VoirBlanc.Checked) ;
             G.Cells[20,Lig]:=VoirBlancs(Enr.DateModif,VoirBlanc.Checked) ;
             G.Cells[21,Lig]:=VoirBlancs(Enr.Soc,VoirBlanc.Checked) ;
             G.Cells[22,Lig]:=VoirBlancs(Enr.Etab,VoirBlanc.Checked) ;
             G.Cells[23,Lig]:=VoirBlancs(Enr.Affaire,VoirBlanc.Checked) ;
             G.Cells[24,Lig]:=VoirBlancs(Enr.DebE,VoirBlanc.Checked) ;
             G.Cells[25,Lig]:=VoirBlancs(Enr.CreE,VoirBlanc.Checked) ;
             G.Cells[26,Lig]:=VoirBlancs(Enr.TauxE,VoirBlanc.Checked) ;
             G.Cells[27,Lig]:=VoirBlancs(Enr.Dev,VoirBlanc.Checked) ;
             G.Cells[28,Lig]:=VoirBlancs(Enr.DebD,VoirBlanc.Checked) ;
             G.Cells[29,Lig]:=VoirBlancs(Enr.CreD,VoirBlanc.Checked) ;
             G.Cells[30,Lig]:=VoirBlancs(Enr.TauxD,VoirBlanc.Checked) ;
             G.Cells[31,Lig]:=VoirBlancs(Enr.DateTauxD,VoirBlanc.Checked) ;
             G.Cells[32,Lig]:=VoirBlancs(Enr.Quot,VoirBlanc.Checked) ;
             G.Cells[33,Lig]:=VoirBlancs(Enr.EcrAN,VoirBlanc.Checked) ;
             G.Cells[34,Lig]:=VoirBlancs(Enr.Qte1,VoirBlanc.Checked) ;
             G.Cells[35,Lig]:=VoirBlancs(Enr.Qte2,VoirBlanc.Checked) ;
             G.Cells[36,Lig]:=VoirBlancs(Enr.Qual1,VoirBlanc.Checked) ;
             G.Cells[37,Lig]:=VoirBlancs(Enr.Qual2,VoirBlanc.Checked) ;
             G.Cells[38,Lig]:=VoirBlancs(Enr.RefLibre,VoirBlanc.Checked) ;
             G.Cells[39,Lig]:=VoirBlancs(Enr.TvaEnc,VoirBlanc.Checked) ;
             G.Cells[40,Lig]:=VoirBlancs(Enr.Regime,VoirBlanc.Checked) ;
             G.Cells[41,Lig]:=VoirBlancs(Enr.Tva,VoirBlanc.Checked) ;
             G.Cells[42,Lig]:=VoirBlancs(Enr.TPF,VoirBlanc.Checked) ;
             G.Cells[43,Lig]:=VoirBlancs(Enr.CtrGen,VoirBlanc.Checked) ;
             G.Cells[44,Lig]:=VoirBlancs(Enr.CtrAux,VoirBlanc.Checked) ;
             G.Cells[45,Lig]:=VoirBlancs(Enr.Couv,VoirBlanc.Checked) ;
             G.Cells[46,Lig]:=VoirBlancs(Enr.Let,VoirBlanc.Checked) ;
             G.Cells[47,Lig]:=VoirBlancs(Enr.LetD,VoirBlanc.Checked) ;
             G.Cells[48,Lig]:=VoirBlancs(Enr.DateMin,VoirBlanc.Checked) ;
             G.Cells[49,Lig]:=VoirBlancs(Enr.DateMax,VoirBlanc.Checked) ;
             G.Cells[50,Lig]:=VoirBlancs(Enr.RefP,VoirBlanc.Checked) ;
             G.Cells[51,Lig]:=VoirBlancs(Enr.DateP,VoirBlanc.Checked) ;
             G.Cells[52,Lig]:=VoirBlancs(Enr.LPLCR,VoirBlanc.Checked) ;
             G.Cells[53,Lig]:=VoirBlancs(Enr.DateRel,VoirBlanc.Checked) ;
             G.Cells[54,Lig]:=VoirBlancs(Enr.Controle,VoirBlanc.Checked) ;
             G.Cells[55,Lig]:=VoirBlancs(Enr.TotEnc,VoirBlanc.Checked) ;
             G.Cells[56,Lig]:=VoirBlancs(Enr.RelEnc,VoirBlanc.Checked) ;
             G.Cells[57,Lig]:=VoirBlancs(Enr.DateV,VoirBlanc.Checked) ;
             G.Cells[58,Lig]:=VoirBlancs(Enr.RIB,VoirBlanc.Checked) ;
             G.Cells[59,Lig]:=VoirBlancs(Enr.RefRel,VoirBlanc.Checked) ;
             G.Cells[60,Lig]:=VoirBlancs(Enr.CouvD,VoirBlanc.Checked) ;
             G.Cells[61,Lig]:=VoirBlancs(Enr.EtatL,VoirBlanc.Checked) ;
             G.Cells[62,Lig]:=VoirBlancs(Enr.NumPInt,VoirBlanc.Checked) ;
             G.Cells[63,Lig]:=VoirBlancs(Enr.Encais,VoirBlanc.Checked) ;
             G.Cells[64,Lig]:=VoirBlancs(Enr.TAN,VoirBlanc.Checked) ;
             G.Cells[65,Lig]:=VoirBlancs(Enr.Eche,VoirBlanc.Checked) ;
             G.Cells[66,Lig]:=VoirBlancs(Enr.Ana,VoirBlanc.Checked) ;
             G.Cells[67,Lig]:=VoirBlancs(Enr.MPaie,VoirBlanc.Checked) ;
             G.Cells[68,Lig]:=VoirBlancs(Enr.NatP,VoirBlanc.Checked) ;
             END ;
     3,4   : BEGIN
             G.Cells[15,Lig]:=VoirBlancs(Enr.Dev,VoirBlanc.Checked) ;
             G.Cells[16,Lig]:=VoirBlancs(Enr.TauxE,VoirBlanc.Checked) ;
             G.Cells[17,Lig]:=VoirBlancs(Enr.CodeMontant,VoirBlanc.Checked) ;
             G.Cells[18,Lig]:=VoirBlancs(Enr.MontantD,VoirBlanc.Checked) ;
             G.Cells[19,Lig]:=VoirBlancs(Enr.MontantE,VoirBlanc.Checked) ;
             G.Cells[20,Lig]:=VoirBlancs(Enr.Etab,VoirBlanc.Checked) ;
             G.Cells[21,Lig]:=VoirBlancs(Enr.Axe,VoirBlanc.Checked) ;
             G.Cells[22,Lig]:=VoirBlancs(Enr.NumEche,VoirBlanc.Checked) ;
             If FormatFic=4 Then
                BEGIN
                G.Cells[23,Lig]:=VoirBlancs(Enr.RefExt,VoirBlanc.Checked) ;
                G.Cells[24,Lig]:=VoirBlancs(Enr.DateRefExt,VoirBlanc.Checked) ;
                G.Cells[25,Lig]:=VoirBlancs(Enr.DateCreat,VoirBlanc.Checked) ;
                G.Cells[26,Lig]:=VoirBlancs(Enr.Soc,VoirBlanc.Checked) ;
                G.Cells[27,Lig]:=VoirBlancs(Enr.Affaire,VoirBlanc.Checked) ;
                G.Cells[28,Lig]:=VoirBlancs(Enr.DateTauxD,VoirBlanc.Checked) ;
                G.Cells[29,Lig]:=VoirBlancs(Enr.EcrAN,VoirBlanc.Checked) ;
                G.Cells[30,Lig]:=VoirBlancs(Enr.Qte1,VoirBlanc.Checked) ;
                G.Cells[31,Lig]:=VoirBlancs(Enr.Qte2,VoirBlanc.Checked) ;
                G.Cells[32,Lig]:=VoirBlancs(Enr.Qual1,VoirBlanc.Checked) ;
                G.Cells[33,Lig]:=VoirBlancs(Enr.Qual2,VoirBlanc.Checked) ;
                G.Cells[34,Lig]:=VoirBlancs(Enr.RefLibre,VoirBlanc.Checked) ;
                G.Cells[35,Lig]:=VoirBlancs(Enr.TvaEnc,VoirBlanc.Checked) ;
                G.Cells[36,Lig]:=VoirBlancs(Enr.Regime,VoirBlanc.Checked) ;
                G.Cells[37,Lig]:=VoirBlancs(Enr.Tva,VoirBlanc.Checked) ;
                G.Cells[38,Lig]:=VoirBlancs(Enr.Tpf,VoirBlanc.Checked) ;
                G.Cells[39,Lig]:=VoirBlancs(Enr.CtrGen,VoirBlanc.Checked) ;
                G.Cells[40,Lig]:=VoirBlancs(Enr.CtrAux,VoirBlanc.Checked) ;
                G.Cells[41,Lig]:=VoirBlancs(Enr.RefP,VoirBlanc.Checked) ;
                G.Cells[42,Lig]:=VoirBlancs(Enr.DateP,VoirBlanc.Checked) ;
                G.Cells[43,Lig]:=VoirBlancs(Enr.DateRel,VoirBlanc.Checked) ;
                G.Cells[44,Lig]:=VoirBlancs(Enr.DateV,VoirBlanc.Checked) ;
                G.Cells[45,Lig]:=VoirBlancs(Enr.RIB,VoirBlanc.Checked) ;
                G.Cells[46,Lig]:=VoirBlancs(Enr.RefRel,VoirBlanc.Checked) ;
                G.Cells[47,Lig]:=VoirBlancs(Enr.Immo,VoirBlanc.Checked) ;
                G.Cells[48,Lig]:=VoirBlancs(Enr.LT0,VoirBlanc.Checked) ;
                G.Cells[49,Lig]:=VoirBlancs(Enr.LT1,VoirBlanc.Checked) ;
                G.Cells[50,Lig]:=VoirBlancs(Enr.LT2,VoirBlanc.Checked) ;
                G.Cells[51,Lig]:=VoirBlancs(Enr.LT3,VoirBlanc.Checked) ;
                G.Cells[52,Lig]:=VoirBlancs(Enr.LT4,VoirBlanc.Checked) ;
                G.Cells[53,Lig]:=VoirBlancs(Enr.LT5,VoirBlanc.Checked) ;
                G.Cells[54,Lig]:=VoirBlancs(Enr.LT6,VoirBlanc.Checked) ;
                G.Cells[55,Lig]:=VoirBlancs(Enr.LT7,VoirBlanc.Checked) ;
                G.Cells[56,Lig]:=VoirBlancs(Enr.LT8,VoirBlanc.Checked) ;
                G.Cells[57,Lig]:=VoirBlancs(Enr.LT9,VoirBlanc.Checked) ;
                G.Cells[58,Lig]:=VoirBlancs(Enr.TA0,VoirBlanc.Checked) ;
                G.Cells[59,Lig]:=VoirBlancs(Enr.TA1,VoirBlanc.Checked) ;
                G.Cells[60,Lig]:=VoirBlancs(Enr.TA2,VoirBlanc.Checked) ;
                G.Cells[61,Lig]:=VoirBlancs(Enr.TA3,VoirBlanc.Checked) ;
                G.Cells[62,Lig]:=VoirBlancs(Enr.LM0,VoirBlanc.Checked) ;
                G.Cells[63,Lig]:=VoirBlancs(Enr.LM1,VoirBlanc.Checked) ;
                G.Cells[64,Lig]:=VoirBlancs(Enr.LM2,VoirBlanc.Checked) ;
                G.Cells[65,Lig]:=VoirBlancs(Enr.LM3,VoirBlanc.Checked) ;
                G.Cells[66,Lig]:=VoirBlancs(Enr.LD,VoirBlanc.Checked) ;
                G.Cells[67,Lig]:=VoirBlancs(Enr.LB0,VoirBlanc.Checked) ;
                G.Cells[68,Lig]:=VoirBlancs(Enr.LB1,VoirBlanc.Checked) ;
                G.Cells[69,Lig]:=VoirBlancs(Enr.Conso,VoirBlanc.Checked) ;
                G.Cells[70,Lig]:=VoirBlancs(Enr.Couv,VoirBlanc.Checked) ;
                G.Cells[71,Lig]:=VoirBlancs(Enr.CouvD,VoirBlanc.Checked) ;
                G.Cells[72,Lig]:=VoirBlancs(Enr.CouvE,VoirBlanc.Checked) ;
                G.Cells[73,Lig]:=VoirBlancs(Enr.DateMax,VoirBlanc.Checked) ;
                G.Cells[74,Lig]:=VoirBlancs(Enr.DateMin,VoirBlanc.Checked) ;
                G.Cells[75,Lig]:=VoirBlancs(Enr.Let,VoirBlanc.Checked) ;
                G.Cells[76,Lig]:=VoirBlancs(Enr.LetD,VoirBlanc.Checked) ;
                G.Cells[77,Lig]:=VoirBlancs(Enr.LetE,VoirBlanc.Checked) ;
                G.Cells[78,Lig]:=VoirBlancs(Enr.EtatL,VoirBlanc.Checked) ;
                END ;
             END ;
     End ;
END ;

Procedure TFCImpFic.RecupEnr(Var Enr : TStLue ; VoirLesBlancs : Boolean ; Lig : Integer) ;
BEGIN
Fillchar(Enr,SizeOf(Enr),#0) ;
Enr.Jal:=VoirBlancs(G.Cells[1,Lig],VoirLesBlancs) ;
Enr.Date:=VoirBlancs(G.Cells[2,Lig],VoirLesBlancs) ;
Enr.TP:=VoirBlancs(G.Cells[3,Lig],VoirLesBlancs) ;
Enr.General:=VoirBlancs(G.Cells[4,Lig],VoirLesBlancs) ;
Enr.TC:=VoirBlancs(G.Cells[5,Lig],VoirLesBlancs) ;
Enr.AuxSect:=VoirBlancs(G.Cells[6,Lig],VoirLesBlancs) ;
Enr.Reference:=VoirBlancs(G.Cells[7,Lig],VoirLesBlancs) ;
Enr.Libelle:=VoirBlancs(G.Cells[8,Lig],VoirLesBlancs) ;
Enr.MP:=VoirBlancs(G.Cells[9,Lig],VoirLesBlancs) ;
Enr.Echeance:=VoirBlancs(G.Cells[10,Lig],VoirLesBlancs) ;
Enr.S:=VoirBlancs(G.Cells[11,Lig],VoirLesBlancs) ;
Enr.Montant:=VoirBlancs(G.Cells[12,Lig],VoirLesBlancs) ;
Enr.TE:=VoirBlancs(G.Cells[13,Lig],VoirLesBlancs) ;
Enr.NumP:=VoirBlancs(G.Cells[14,Lig],VoirLesBlancs) ;
Case FormatFic Of
  2 : BEGIN { HALLEY ETENDU }
      Enr.NumEche:=VoirBlancs(G.Cells[15,Lig],VoirLesBlancs) ;
      Enr.Valid:=VoirBlancs(G.Cells[16,Lig],VoirLesBlancs) ;
      Enr.RefExt:=VoirBlancs(G.Cells[17,Lig],VoirLesBlancs) ;
      Enr.DateRefExt:=VoirBlancs(G.Cells[18,Lig],VoirLesBlancs) ;
      Enr.DateCreat:=VoirBlancs(G.Cells[19,Lig],VoirLesBlancs) ;
      Enr.DateModif:=VoirBlancs(G.Cells[20,Lig],VoirLesBlancs) ;
      Enr.Soc:=VoirBlancs(G.Cells[21,Lig],VoirLesBlancs) ;
      Enr.Etab:=VoirBlancs(G.Cells[22,Lig],VoirLesBlancs) ;
      Enr.Affaire:=VoirBlancs(G.Cells[23,Lig],VoirLesBlancs) ;
      Enr.DebE:=VoirBlancs(G.Cells[24,Lig],VoirLesBlancs) ;
      Enr.CreE:=VoirBlancs(G.Cells[25,Lig],VoirLesBlancs) ;
      Enr.TauxE:=VoirBlancs(G.Cells[26,Lig],VoirLesBlancs) ;
      Enr.Dev:=VoirBlancs(G.Cells[27,Lig],VoirLesBlancs) ;
      Enr.DebD:=VoirBlancs(G.Cells[28,Lig],VoirLesBlancs) ;
      Enr.CreD:=VoirBlancs(G.Cells[29,Lig],VoirLesBlancs) ;
      Enr.TauxD:=VoirBlancs(G.Cells[30,Lig],VoirLesBlancs) ;
      Enr.DateTauxD:=VoirBlancs(G.Cells[31,Lig],VoirLesBlancs) ;
      Enr.Quot:=VoirBlancs(G.Cells[32,Lig],VoirLesBlancs) ;
      Enr.EcrAN:=VoirBlancs(G.Cells[33,Lig],VoirLesBlancs) ;
      Enr.Qte1:=VoirBlancs(G.Cells[34,Lig],VoirLesBlancs) ;
      Enr.Qte2:=VoirBlancs(G.Cells[35,Lig],VoirLesBlancs) ;
      Enr.Qual1:=VoirBlancs(G.Cells[36,Lig],VoirLesBlancs) ;
      Enr.Qual2:=VoirBlancs(G.Cells[37,Lig],VoirLesBlancs) ;
      Enr.RefLibre:=VoirBlancs(G.Cells[38,Lig],VoirLesBlancs) ;
      Enr.TvaEnc:=VoirBlancs(G.Cells[39,Lig],VoirLesBlancs) ;
      Enr.Regime:=VoirBlancs(G.Cells[40,Lig],VoirLesBlancs) ;
      Enr.Tva:=VoirBlancs(G.Cells[41,Lig],VoirLesBlancs) ;
      Enr.TPF:=VoirBlancs(G.Cells[42,Lig],VoirLesBlancs) ;
      Enr.CtrGen:=VoirBlancs(G.Cells[43,Lig],VoirLesBlancs) ;
      Enr.CtrAux:=VoirBlancs(G.Cells[44,Lig],VoirLesBlancs) ;
      Enr.Couv:=VoirBlancs(G.Cells[45,Lig],VoirLesBlancs) ;
      Enr.Let:=VoirBlancs(G.Cells[46,Lig],VoirLesBlancs) ;
      Enr.LetD:=VoirBlancs(G.Cells[47,Lig],VoirLesBlancs) ;
      Enr.DateMin:=VoirBlancs(G.Cells[48,Lig],VoirLesBlancs) ;
      Enr.DateMax:=VoirBlancs(G.Cells[49,Lig],VoirLesBlancs) ;
      Enr.RefP:=VoirBlancs(G.Cells[50,Lig],VoirLesBlancs) ;
      Enr.DateP:=VoirBlancs(G.Cells[51,Lig],VoirLesBlancs) ;
      Enr.LPLCR:=VoirBlancs(G.Cells[52,Lig],VoirLesBlancs) ;
      Enr.DateRel:=VoirBlancs(G.Cells[53,Lig],VoirLesBlancs) ;
      Enr.Controle:=VoirBlancs(G.Cells[54,Lig],VoirLesBlancs) ;
      Enr.TotEnc:=VoirBlancs(G.Cells[55,Lig],VoirLesBlancs) ;
      Enr.RelEnc:=VoirBlancs(G.Cells[56,Lig],VoirLesBlancs) ;
      Enr.DateV:=VoirBlancs(G.Cells[57,Lig],VoirLesBlancs) ;
      Enr.RIB:=VoirBlancs(G.Cells[58,Lig],VoirLesBlancs) ;
      Enr.RefRel:=VoirBlancs(G.Cells[59,Lig],VoirLesBlancs) ;
      Enr.CouvD:=VoirBlancs(G.Cells[60,Lig],VoirLesBlancs) ;
      Enr.EtatL:=VoirBlancs(G.Cells[61,Lig],VoirLesBlancs) ;
      Enr.NumPInt:=VoirBlancs(G.Cells[62,Lig],VoirLesBlancs) ;
      Enr.Encais:=VoirBlancs(G.Cells[63,Lig],VoirLesBlancs) ;
      Enr.TAN:=VoirBlancs(G.Cells[64,Lig],VoirLesBlancs) ;
      Enr.Eche:=VoirBlancs(G.Cells[65,Lig],VoirLesBlancs) ;
      Enr.Ana:=VoirBlancs(G.Cells[66,Lig],VoirLesBlancs) ;
      Enr.MPaie:=VoirBlancs(G.Cells[67,Lig],VoirLesBlancs) ;
      Enr.NatP:=VoirBlancs(G.Cells[68,Lig],VoirLesBlancs) ;
      END ;
  END ;
END ;

Function TFCImpFic.GereTotalpiece(WhatErreur : Integer ; OldEnr,Enr : TStLue ; Var TotPieceD,TotPieceC : Double ;
                                   Var NumP,NumPE,NumPS : Integer) : String ;
Var Montant : Double ;
    StPb,StSauveNum : String ;
BEGIN
StPb:='' ; result:='' ;
If WhatErreur in [1]=FALSE Then
   BEGIN
   If (WhatErreur In [7,8])=FALSE Then Montant:=StrToFloat(StPoint(Enr.Montant)) Else Montant:=0 ;
   If Enr.NumP<>OldEnr.NumP Then
      BEGIN
      StSauveNum:='' ;
      If Pos('/',GP.Cells[0,GP.RowCount-1])>0 Then StSauveNum:=GP.Cells[0,GP.RowCount-1] ;
      GP.Cells[0,GP.RowCount-1]:=OldEnr.NumP ;
      If Arrondi(TotPieceD-TotPieceC,V_PGI.OkDecV)<>0 Then
         BEGIN
         StNumP:=StNumP+Trim(OldEnr.NumP)+';' ;
         GP.Cells[0,GP.RowCount-1]:='*'+OldEnr.NumP+'*' ;
         Inc(NumPE) ;
         END Else
         BEGIN
         If StSauveNum<>'' Then GP.Cells[0,GP.RowCount-1]:=StSauveNum ;
         END ;
      Inc(NumP) ;
      GP.Cells[1,GP.RowCount-1]:=OldEnr.Jal ;
      GP.Cells[2,GP.RowCount-1]:=OldEnr.Date ;
      GP.Cells[3,GP.RowCount-1]:=FormatFloat('#,##0.00',TotPieceD) ;
      GP.Cells[4,GP.RowCount-1]:=FormatFloat('#,##0.00',TotPieceC) ;
      GP.Cells[5,GP.RowCount-1]:=FormatFloat('#,##0.00',TotPieceD-TotPieceC) ;
      GP.RowCount:=GP.RowCount+1 ;
      TotPieceD:=0 ; TotPieceC:=0 ;
      END Else
      BEGIN
      If Enr.Jal<>OldEnr.Jal Then StPb:=StPb+'J' ;
      If Enr.Date<>OldEnr.Date Then StPb:=StPb+'D' ;
      If Enr.TP<>OldEnr.TP Then StPb:=StPb+'P' ;
      If Enr.TE<>OldEnr.TE Then StPb:=StPb+'T' ;
      If FormatFic in [3,4] Then
         BEGIN
         If Enr.Dev<>OldEnr.Dev Then StPb:=StPb+'$' ;
         If Enr.TauxD<>OldEnr.TauxD Then StPb:=StPb+'X' ;
         END ;
      Case FormatFic Of
        2,3,4 : BEGIN
                If Enr.Etab<>OldEnr.Etab Then StPb:=StPb+'E' ;
                END ;
        END ;
      If StPb<>'' Then
         BEGIN
         Result:=StPb ;
         If Pos('/',GP.Cells[0,GP.RowCount-1])=0 Then GP.Cells[0,GP.RowCount-1]:='*'+OldEnr.NumP+'*/'+StPb ;
         Inc(NumPS) ;
         StNumP:=StNumP+Trim(OldEnr.NumP)+';' ;
         END ;
      END ;
   if (Enr.TC<>'A') And (Enr.TC<>'E') And (Enr.TC<>'D') then  { pour shunter les écritures analytiques }
      BEGIN
      If Enr.S='D' Then TotPieceD:=Arrondi(TotPieceD+Montant,V_PGI.OkDecV)
                   Else TotPieceC:=Arrondi(TotPieceC+Montant,V_PGI.OkDecV) ;
      END ;
   END ;
END ;

function AlimLeCptLu(Enr : TStLue ; Cpt : String ; Var TabCpt : TTabCptLu) : TCptLu ;
Var i : Integer ;
    LeCptLu : TCptLu ;
BEGIN
Fillchar(LeCptLu,SizeOf(LeCptLu),#0) ; Result:=LeCptLu ;
For i:=1 To MaxCptLu Do
  BEGIN
  If TabCpt[i].Cpt='' Then Break ;
  If Cpt=TabCpt[i].Cpt Then BEGIN Result:=TabCpt[i] ; Exit ; END ;
  END ;
TabCpt[i].Cpt:=Cpt ;
TabCpt[i].Collectif:=Enr.General ;
Result:=TabCpt[i] ;
END ;

Procedure RetoucheCptLu(LeCptLu : TCptLu ; Var TabCpt : TTabCptLu) ;
Var i : Integer ;
BEGIN
If LeCptLu.Collectif='' Then Exit ;
For i:=1 To MaxCptLu Do
  BEGIN
  If TabCpt[i].Cpt='' Then Break ;
  If LeCptLu.Collectif=TabCpt[i].Cpt Then Break ;
  END ;
TabCpt[i].VEntilable:=TRUE ;
END ;

Procedure TFCImpFic.AlimCptLu(Enr : TStLue) ;
Var Cpt : String ;
    LeCptLu : TCptLu ;
BEGIN
Cpt:=Trim(Enr.General) ; AlimLeCptLu(Enr,Cpt,GenLu) ;
Cpt:=Trim(Enr.AuxSect) ;
If Cpt<>'' Then
  Case Enr.TC[1] Of
    'X' : AlimLeCptLu(Enr,Cpt,AuxLu) ;
    'A' : BEGIN
          LeCptLu:=AlimLeCptLu(Enr,Cpt,SecLu) ;
          LeCptLu.Collectif:=Trim(Enr.General) ;
          RetoucheCptLu(LeCptLu,GenLu) ;
          END ;
    END ;
END ;

Function TFCImpFic.AlimGridCpt(What : Integer ; Var St : String ; Lig : Integer) : Boolean ;
Var Qui : String ;
BEGIN
Result:=FALSE ; Qui:=Copy(St,4,3) ;
If (Qui<>'CGE') And (Qui<>'CAU') And (Qui<>'CAS') And (Qui<>'CAE') And (Qui<>'SAN')
{AJOUT ME 09-10-01}
AND (Qui <>'PS1') AND (Qui <>'PS2') AND (Qui <>'PS3') AND (Qui <>'PS4') AND (Qui <>'PS5')
AND (Qui <>'EXO')Then Exit ;

If ((Qui='CAU') Or (Qui='CAS') Or (Qui='CAE')) And (What<>1) Then Exit ;
If (Qui='CGE') And (What<>0) Then Exit ;
If (Qui='SAN') And (What<>2) Then Exit ;
{AJOUT ME 09-10-01}
if ((Qui ='PS1') OR (Qui ='PS2') OR (Qui ='PS3') OR (Qui ='PS4') OR (Qui ='PS5')) AND (What <> 3) Then Exit ;
// ajout me 13-12-01
if ((Qui ='EXO')) AND (What <> 4) Then Exit ;


If Qui='CGE' Then
  BEGIN
  GGen.Cells[0,Lig]:=Copy(St,1,3) ;
  GGen.Cells[1,Lig]:=Copy(St,4,3) ;
  GGen.Cells[2,Lig]:=Copy(St,7,17) ;
  GGen.Cells[3,Lig]:=Copy(St,24,35) ;
  GGen.Cells[4,Lig]:=Copy(St,59,3) ;
  GGen.Cells[5,Lig]:=Copy(St,62,1) ;
  GGen.Cells[6,Lig]:=Copy(St,63,1) ;
  GGen.Cells[7,Lig]:=Copy(St,64,1) ;
  GGen.Cells[8,Lig]:=Copy(St,65,1) ;
  GGen.Cells[9,Lig]:=Copy(St,66,1) ;
  GGen.Cells[10,Lig]:=Copy(St,67,1) ;
  GGen.Cells[11,Lig]:=Copy(St,68,1) ;
  GGen.Cells[12,Lig]:=Copy(St,69,3) ;
  GGen.Cells[13,Lig]:=Copy(St,72,3) ;
  GGen.Cells[14,Lig]:=Copy(St,75,3) ;
  GGen.Cells[15,Lig]:=Copy(St,78,3) ;
  GGen.Cells[16,Lig]:=Copy(St,81,3) ;
  GGen.Cells[17,Lig]:=Copy(St,84,3) ;
  GGen.Cells[18,Lig]:=Copy(St,87,3) ;
  GGen.Cells[19,Lig]:=Copy(St,90,3) ;
  GGen.Cells[20,Lig]:=Copy(St,93,3) ;
  GGen.Cells[21,Lig]:=Copy(St,96,3) ;
  GGen.Cells[22,Lig]:=Copy(St,99,17) ;
  GGen.Cells[23,Lig]:=Copy(St,116,3) ;
  Result:=TRUE ;
  END Else
If (Qui='CAU') Or (Qui='CAS') Then
  BEGIN
  GAux.Cells[0,Lig]:=Copy(St,1,3) ;
  GAux.Cells[1,Lig]:=Copy(St,4,3) ;
  GAux.Cells[2,Lig]:=Copy(St,7,17) ;
  GAux.Cells[3,Lig]:=Copy(St,24,35) ;
  GAux.Cells[4,Lig]:=Copy(St,59,3) ;
  GAux.Cells[5,Lig]:=Copy(St,62,1) ;
  GAux.Cells[6,Lig]:=Copy(St,63,17) ;
  GAux.Cells[7,Lig]:=Copy(St,80,17) ;
  If Qui='CAU' Then
    BEGIN
    GAux.Cells[8,Lig]:=Copy(St,97,3) ;
    GAux.Cells[9,Lig]:=Copy(St,100,3) ;
    GAux.Cells[10,Lig]:=Copy(St,103,3) ;
    GAux.Cells[11,Lig]:=Copy(St,106,3) ;
    GAux.Cells[12,Lig]:=Copy(St,109,3) ;
    GAux.Cells[13,Lig]:=Copy(St,112,3) ;
    GAux.Cells[14,Lig]:=Copy(St,115,3) ;
    GAux.Cells[15,Lig]:=Copy(St,118,3) ;
    GAux.Cells[16,Lig]:=Copy(St,121,3) ;
    GAux.Cells[17,Lig]:=Copy(St,124,3) ;
    END Else
    BEGIN
    GAux.Cells[8,Lig]:=Copy(St,97,4) ;
    GAux.Cells[9,Lig]:=Copy(St,101,4) ;
    GAux.Cells[10,Lig]:=Copy(St,105,4) ;
    GAux.Cells[11,Lig]:=Copy(St,109,4) ;
    GAux.Cells[12,Lig]:=Copy(St,113,2) ;
    GAux.Cells[13,Lig]:=Copy(St,115,2) ;
    GAux.Cells[14,Lig]:=Copy(St,117,2) ;
    GAux.Cells[15,Lig]:=Copy(St,119,2) ;
    GAux.Cells[16,Lig]:=Copy(St,121,3) ;
    GAux.Cells[17,Lig]:=Copy(St,124,3) ;
    END ;
  GAux.Cells[18,Lig]:=Copy(St,127,35) ;
  GAux.Cells[19,Lig]:=Copy(St,162,35) ;
  GAux.Cells[20,Lig]:=Copy(St,197,35) ;
  GAux.Cells[21,Lig]:=Copy(St,232,9) ;
  GAux.Cells[22,Lig]:=Copy(St,241,35) ;
  GAux.Cells[23,Lig]:=Copy(St,276,24) ;
  GAux.Cells[24,Lig]:=Copy(St,300,5) ;
  GAux.Cells[25,Lig]:=Copy(St,305,5) ;
  GAux.Cells[26,Lig]:=Copy(St,310,11) ;
  GAux.Cells[27,Lig]:=Copy(St,321,2) ;
  GAux.Cells[28,Lig]:=Copy(St,323,3) ;
  GAux.Cells[29,Lig]:=Copy(St,326,17) ;
  GAux.Cells[30,Lig]:=Copy(St,343,3) ;
  GAux.Cells[31,Lig]:=Copy(St,346,1) ;
  GAux.Cells[32,Lig]:=Copy(St,347,3) ;
  GAux.Cells[33,Lig]:=Copy(St,350,25) ;
  GAux.Cells[34,Lig]:=Copy(St,375,25) ;
  GAux.Cells[35,Lig]:=Copy(St,400,3) ;
  GAux.Cells[36,Lig]:=Copy(St,403,3) ;
  GAux.Cells[37,Lig]:=Copy(St,406,35) ;
  Result:=TRUE ;
  END Else
If (Qui='CAE') Then
  BEGIN
  GAux.Cells[0,Lig]:=Copy(St,1,3) ;
  GAux.Cells[1,Lig]:=Copy(St,4,3) ;
  GAux.Cells[2,Lig]:=Copy(St,7,17) ;
  GAux.Cells[3,Lig]:=Copy(St,24,35) ;
  GAux.Cells[4,Lig]:=Copy(St,59,3) ;
  GAux.Cells[5,Lig]:=Copy(St,62,1) ;
  GAux.Cells[6,Lig]:=Copy(St,63,17) ;
  GAux.Cells[7,Lig]:=Copy(St,80,17) ;
  GAux.Cells[8,Lig]:=Copy(St,97,17) ;
  GAux.Cells[9,Lig]:=Copy(St,114,17) ;
  GAux.Cells[10,Lig]:=Copy(St,131,17) ;
  GAux.Cells[11,Lig]:=Copy(St,148,17) ;
  GAux.Cells[12,Lig]:=Copy(St,165,17) ;
  GAux.Cells[13,Lig]:=Copy(St,182,17) ;
  GAux.Cells[14,Lig]:=Copy(St,199,17) ;
  GAux.Cells[15,Lig]:=Copy(St,216,17) ;
  GAux.Cells[16,Lig]:=Copy(St,233,17) ;
  GAux.Cells[17,Lig]:=Copy(St,250,17) ;
  GAux.Cells[18,Lig]:=Copy(St,267,35) ;
  GAux.Cells[19,Lig]:=Copy(St,302,35) ;
  GAux.Cells[20,Lig]:=Copy(St,337,35) ;
  GAux.Cells[21,Lig]:=Copy(St,372,9) ;
  GAux.Cells[22,Lig]:=Copy(St,381,35) ;
  GAux.Cells[23,Lig]:=Copy(St,416,24) ;
  GAux.Cells[24,Lig]:=Copy(St,440,5) ;
  GAux.Cells[25,Lig]:=Copy(St,445,5) ;
  GAux.Cells[26,Lig]:=Copy(St,450,11) ;
  GAux.Cells[27,Lig]:=Copy(St,461,2) ;
  GAux.Cells[28,Lig]:=Copy(St,463,3) ;
  GAux.Cells[29,Lig]:=Copy(St,466,17) ;
  GAux.Cells[30,Lig]:=Copy(St,483,3) ;
  GAux.Cells[31,Lig]:=Copy(St,486,1) ;
  GAux.Cells[32,Lig]:=Copy(St,487,3) ;
  GAux.Cells[33,Lig]:=Copy(St,490,25) ;
  GAux.Cells[34,Lig]:=Copy(St,515,25) ;
  GAux.Cells[35,Lig]:=Copy(St,540,3) ;
  GAux.Cells[36,Lig]:=Copy(St,543,3) ;
  GAux.Cells[37,Lig]:=Copy(St,546,35) ;
  Result:=TRUE ;
  END Else
If Qui='SAN' Then
  BEGIN
  GSect.Cells[0,Lig]:=Copy(St,1,3) ;
  GSect.Cells[1,Lig]:=Copy(St,4,3) ;
  GSect.Cells[2,Lig]:=Copy(St,7,17) ;
  GSect.Cells[3,Lig]:=Copy(St,24,35) ;
  GSect.Cells[4,Lig]:=Copy(St,59,3) ;
  GSect.Cells[5,Lig]:=Copy(St,62,3) ;
  GSect.Cells[6,Lig]:=Copy(St,65,3) ;
  GSect.Cells[7,Lig]:=Copy(St,68,3) ;
  GSect.Cells[8,Lig]:=Copy(St,71,3) ;
  GSect.Cells[9,Lig]:=Copy(St,74,3) ;
  GSect.Cells[10,Lig]:=Copy(St,77,3) ;
  GSect.Cells[11,Lig]:=Copy(St,80,3) ;
  GSect.Cells[12,Lig]:=Copy(St,83,3) ;
  GSect.Cells[13,Lig]:=Copy(St,86,3) ;
  GSect.Cells[14,Lig]:=Copy(St,89,3) ;
  GSect.Cells[15,Lig]:=Copy(St,92,17) ;
  GSect.Cells[16,Lig]:=Copy(St,109,3) ;
  Result:=TRUE ;
  END ;

{AJOUT ME 09-10-01 paramétres généraux1}
If Qui='PS1' Then
begin
  G7.Cells[0,Lig]:=Copy(St,1,3) ;
  G7.Cells[1,Lig]:=Copy(St,4,3) ;
  G7.Cells[2,Lig]:=Copy(St,7,35) ;
  G7.Cells[3,Lig]:=Copy(St,42,35) ;
  G7.Cells[4,Lig]:=Copy(St,77,35) ;
  G7.Cells[5,Lig]:=Copy(St,112,35) ;
  G7.Cells[6,Lig]:=Copy(St,147,9) ;
  G7.Cells[7,Lig]:=Copy(St,156,35) ;
  G7.Cells[8,Lig]:=Copy(St,191,3) ;
  G7.Cells[9,Lig]:=Copy(St,194,25) ;
  G7.Cells[10,Lig]:=Copy(St,219,25) ;
  G7.Cells[11,Lig]:=Copy(St,244,25) ;
  G7.Cells[12,Lig]:=Copy(St,269,35) ;
  G7.Cells[13,Lig]:=Copy(St,304,35) ;
  G7.Cells[14,Lig]:=Copy(St,339,35) ;
  G7.Cells[15,Lig]:=Copy(St,374,17) ;
  G7.Cells[16,Lig]:=Copy(St,391,35) ;
  G7.Cells[17,Lig]:=Copy(St,426,17) ;
  G7.Cells[16,Lig]:=Copy(St,443,35) ;
  G7.Cells[17,Lig]:=Copy(St,478,20) ;
  G7.Cells[18,Lig]:=Copy(St,498,35) ;
  Result:=TRUE ;
end;
{AJOUT ME 09-10-01 paramétres généraux2}
If Qui='PS2' Then
begin
  G8.Cells[0,Lig]:=Copy(St,1,3) ;
  G8.Cells[1,Lig]:=Copy(St,4,3) ;
  G8.Cells[2,Lig]:=Copy(St,7,2) ;
  G8.Cells[3,Lig]:=Copy(St,9,1) ;
  G8.Cells[4,Lig]:=Copy(St,10,2) ;
  G8.Cells[5,Lig]:=Copy(St,12,1) ;
  G8.Cells[6,Lig]:=Copy(St,13,2) ;
  G8.Cells[7,Lig]:=Copy(St,15,1) ;
  G8.Cells[8,Lig]:=Copy(St,16,2) ;
  G8.Cells[9,Lig]:=Copy(St,18,1) ;
  G8.Cells[10,Lig]:=Copy(St,19,2) ;
  G8.Cells[11,Lig]:=Copy(St,21,1) ;
  G8.Cells[12,Lig]:=Copy(St,22,2) ;
  G8.Cells[13,Lig]:=Copy(St,24,1) ;
  G8.Cells[14,Lig]:=Copy(St,25,2) ;
  G8.Cells[15,Lig]:=Copy(St,27,1) ;
  G8.Cells[16,Lig]:=Copy(St,28,17) ;
  G8.Cells[17,Lig]:=Copy(St,45,17) ;
  G8.Cells[18,Lig]:=Copy(St,62,17) ;
  G8.Cells[19,Lig]:=Copy(St,79,17) ;
  G8.Cells[20,Lig]:=Copy(St,96,17) ;
  G8.Cells[21,Lig]:=Copy(St,113,17) ;
  G8.Cells[22,Lig]:=Copy(St,130,17) ;
  G8.Cells[23,Lig]:=Copy(St,147,17) ;
  G8.Cells[24,Lig]:=Copy(St,164,17) ;
  G8.Cells[25,Lig]:=Copy(St,181,17) ;
  Result:=TRUE ;
end;

// paramétres généraux3
If Qui='PS3' Then
begin
  G9.Cells[0,Lig]:=Copy(St,1,3) ;
  G9.Cells[1,Lig]:=Copy(St,4,3) ;
  G9.Cells[2,Lig]:=Copy(St,7,17) ;
  G9.Cells[3,Lig]:=Copy(St,24,17) ;
  G9.Cells[4,Lig]:=Copy(St,41,17) ;
  G9.Cells[5,Lig]:=Copy(St,58,17) ;
  G9.Cells[6,Lig]:=Copy(St,75,3) ;
  G9.Cells[7,Lig]:=Copy(St,78,3) ;
  G9.Cells[8,Lig]:=Copy(St,81,3) ;
  Result:=TRUE ;
end;

// paramétres généraux4
If Qui='PS4' Then
begin
  G10.Cells[0,Lig]:=Copy(St,1,3) ;
  G10.Cells[1,Lig]:=Copy(St,4,3) ;
  G10.Cells[2,Lig]:=Copy(St,7,17) ;
  G10.Cells[3,Lig]:=Copy(St,24,17) ;
  G10.Cells[4,Lig]:=Copy(St,41,17) ;
  G10.Cells[5,Lig]:=Copy(St,58,17) ;
  G10.Cells[6,Lig]:=Copy(St,75,17) ;
  G10.Cells[7,Lig]:=Copy(St,92,17) ;
  G10.Cells[8,Lig]:=Copy(St,109,17) ;
  G10.Cells[9,Lig]:=Copy(St,126,17) ;
  G10.Cells[10,Lig]:=Copy(St,143,17) ;
  G10.Cells[11,Lig]:=Copy(St,160,17) ;
  G10.Cells[12,Lig]:=Copy(St,177,17) ;
  G10.Cells[13,Lig]:=Copy(St,194,17) ;
  G10.Cells[14,Lig]:=Copy(St,211,17) ;
  G10.Cells[15,Lig]:=Copy(St,228,17) ;
  G10.Cells[16,Lig]:=Copy(St,245,17) ;
  G10.Cells[17,Lig]:=Copy(St,262,17) ;
  G10.Cells[18,Lig]:=Copy(St,279,17) ;
  G10.Cells[19,Lig]:=Copy(St,296,17) ;
  G10.Cells[20,Lig]:=Copy(St,313,17) ;
  G10.Cells[21,Lig]:=Copy(St,330,17) ;
  G10.Cells[22,Lig]:=Copy(St,347,17) ;
  G10.Cells[23,Lig]:=Copy(St,364,17) ;
  G10.Cells[24,Lig]:=Copy(St,381,17) ;
  G10.Cells[25,Lig]:=Copy(St,398,17) ;
  G10.Cells[26,Lig]:=Copy(St,415,17) ;
  G10.Cells[27,Lig]:=Copy(St,432,17) ;
  G10.Cells[28,Lig]:=Copy(St,449,17) ;
  G10.Cells[29,Lig]:=Copy(St,466,17) ;
  G10.Cells[30,Lig]:=Copy(St,483,17) ;
  G10.Cells[31,Lig]:=Copy(St,500,17) ;
  Result:=TRUE ;
end;
If Qui='PS5' Then
begin
  G11.Cells[0,Lig]:=Copy(St,1,3) ;
  G11.Cells[1,Lig]:=Copy(St,4,3) ;
  G11.Cells[2,Lig]:=Copy(St,7,3) ;
  G11.Cells[3,Lig]:=Copy(St,10,1) ;
  G11.Cells[4,Lig]:=Copy(St,11,1) ;
  G11.Cells[5,Lig]:=Copy(St,12,3) ;
  G11.Cells[6,Lig]:=Copy(St,15,17) ;
  G11.Cells[7,Lig]:=Copy(St,32,17) ;
  G11.Cells[8,Lig]:=Copy(St,49,10) ;
  G11.Cells[9,Lig]:=Copy(St,59,3) ;
  G11.Cells[10,Lig]:=Copy(St,62,3) ;
  G11.Cells[11,Lig]:=Copy(St,65,3) ;
  G11.Cells[12,Lig]:=Copy(St,68,3) ;
end;
If Qui='EXO' Then
begin
  G12.Cells[0,Lig]:=Copy(St,1,3) ;
  G12.Cells[1,Lig]:=Copy(St,4,3) ;
  G12.Cells[2,Lig]:=Copy(St,7,3) ;
  G12.Cells[3,Lig]:=Copy(St,10,8) ;
  G12.Cells[4,Lig]:=Copy(St,18,8) ;
  G12.Cells[5,Lig]:=Copy(St,26,3) ;
  G12.Cells[6,Lig]:=Copy(St,29,3) ;
  G12.Cells[7,Lig]:=Copy(St,32,35) ;
  // ajout me 13-12-01
  Result:=TRUE ;
end;

END ;

Function TFCImpFic.ChargeGrid : Integer ;
Var FF   : TextFile ;
    St,StPb : String ;
    Okok,UneErreur,OkSt : boolean ;
    NumL,NumLE,WhatErreur,NumP,NumPE,NumPS : Integer ;
    TotPieceD,TotPieceC : Double ;
    LigneCpt : Boolean ;
    OldEnr,Enr : TStLue ;
BEGIN
Result:=0 ; Fillchar(GenLu,SizeOf(GenLu),#0) ; Fillchar(AuxLu,SizeOf(AuxLu),#0) ;
Fillchar(SecLu,SizeOf(SecLu),#0) ;
AssignFile(FF,StFichier) ; Okok:=False ; StNumP:='' ;
{$I-} Reset(FF) ; {$I+} if IoResult<>0 then Exit ;
//G.VidePile(False) ;
NumL:=1 ; NumLE:=0 ; NumP:=0 ; NumPE:=0 ; NumPS:=0 ;
TotPieceD:=0 ; TotPieceC:=0 ; WhatErreur:=0 ;
try
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
While Not EOF(FF) do
   BEGIN
   MoveCur(FALSE) ;
   Readln(FF,St) ;
   LigneCpt:=EstUneLigneCpt(St) ;
   OkSt:=(Trim(St)<>'') And (Not LigneCpt) ;
   If OkSt Then
      BEGIN
      OkOk:=TRUE ;
      UneErreur:=FALSE ;
      AlimEnr(St,Enr,FormatFic) ; Enr.NumLg:=IntToStr(NumL) ; WhatErreur:=0 ;
      If NumL>1 Then BEGIN WhatErreur:=VerifEnr(Enr) ; UneErreur:=WhatErreur>0 ; END ;
      If NumL=2 Then OldEnr:=Enr ;
      AlimGrid(Enr,G.RowCount-1,FormatFic) ;
      If NumL>1 Then AlimCptLu(Enr) ;
      StPb:='' ;
      If NumL>1 Then StPb:=GereTotalpiece(WhatErreur,OldEnr,Enr,TotPieceD,TotPieceC,NumP,NumPE,NumPS) ;
      If Not UneErreur Then UneErreur:=StPb<>'' ;
      if UneErreur then
         BEGIN
         G.Cells[0,G.RowCount-1]:='*'+IntToStr(NumL)+'*' ; Inc(NumLE) ;
         If StPb<>'' Then G.Cells[0,G.RowCount-1]:=G.Cells[0,G.RowCount-1]+'/'+StPb ;
         END ;
      G.RowCount:=G.RowCount+1 ; Inc(NumL) ; OldEnr:=Enr ;
      END Else If LigneCpt And (FormatFic in [3,4]) And (Trim(St)<>'') Then
      BEGIN
      If AlimGridCpt(0,St,GGen.RowCount-1) Then GGen.RowCount:=GGen.RowCount+1 ;
      If AlimGridCpt(1,St,GAux.RowCount-1) Then GAux.RowCount:=GAux.RowCount+1 ;
      If AlimGridCpt(2,St,GSect.RowCount-1) Then GSect.RowCount:=GSect.RowCount+1 ;
// pour les généraux   AJOUT ME 09-10-01
      AlimGridCpt(3,St,1);
      // ajout me 13-12-01  pour les exercices
      If AlimGridCpt(4,St,G12.RowCount-1) Then G12.RowCount:=G12.RowCount+1 ;
      END ;
   END ;
If NumL>2 Then
   BEGIN
   Enr.NumP:='' ; GereTotalpiece(WhatErreur,OldEnr,Enr,TotPieceD,TotPieceC,NumP,NumPE,NumPS) ;
   END ;
EXCEPT
    On E: Exception do
      begin
      ShowMessage ( 'TozError : ' + E.Message ) ;
      end ;
END ;

CloseFile(FF) ;
FiniMove ;
if Okok then BEGIN G.RowCount:=G.RowCount-1 ; GP.RowCount:=GP.RowCount-1 ;END ;
Dec(NumL) ;
NbLignes.Caption:=NbLignes.Caption+' '+IntToStr(NumL) ;
NbLignesFausses.Caption:=NbLignesFausses.Caption+' '+IntToStr(NumLE) ;
Nbpieces.Caption:=NbPieces.Caption+' '+IntToStr(NumP) ;
NbPiecesFausses.Caption:=NbPiecesFausses.Caption+' '+IntToStr(NumPE) ;
NbPiecesIncorrectes.Caption:=NbPiecesIncorrectes.Caption+' '+IntToStr(NumPS) ;
Result:=NumLE ;
VisuAux.Enabled:=GAux.RowCount>2 ;
VisuGen.Enabled:=GGen.RowCount>2 ;
VisuSect.Enabled:=GSect.RowCount>2 ;
END ;

procedure TFCImpFic.InitGrid ;
Var i : Integer ;
BEGIN
Case FormatFic Of
  0,1 : BEGIN
        For i:=0 To 14 Do
           BEGIN
           G.Titres[i]:=HFormatSaari.Mess[i] ; GPL.Titres[i]:=HFormatSaari.Mess[i] ;
           END ;
        If sn2Orli Then
           BEGIN
           G.Titres[15]:='Mtant. Devise' ;
           G.Titres[16]:='Dev.' ;
           G.Titres[17]:='Taux' ;
           G.Titres[18]:='MP' ;
           G.Titres[19]:='RIB' ;
           For i:=20 To G.ColCount-1 Do
              BEGIN
              G.ColWidths[i]:=0 ; GPL.ColWidths[i]:=0 ;
              END ;
           END Else
           BEGIN
           For i:=15 To G.ColCount-1 Do
              BEGIN
              G.ColWidths[i]:=0 ; GPL.ColWidths[i]:=0 ;
              END ;
           END ;
        END ;
  2   : BEGIN
        For i:=0 To 68 Do
           BEGIN
           G.Titres[i]:=HFormatHALLEY.Mess[i] ; GPL.Titres[i]:=HFormatHALLEY.Mess[i] ;
           END ;
        For i:=69 To G.ColCount-1 Do
           BEGIN
           G.ColWidths[i]:=0 ; GPL.ColWidths[i]:=0 ;
           END ;
        END ;
  3   : BEGIN
        For i:=0 To 21 Do
           BEGIN
           G.Titres[i]:=HFormatCGN.Mess[i] ; GPL.Titres[i]:=HFormatCGN.Mess[i] ;
           END ;
        For i:=23 To G.ColCount-1 Do
           BEGIN
           G.ColWidths[i]:=0 ; GPL.ColWidths[i]:=0 ;
           END ;
        END ;
  4   : BEGIN
        For i:=0 To G.ColCount-1 Do
           BEGIN
           G.Titres[i]:=HFormatCGE.Mess[i] ; GPL.Titres[i]:=HFormatCGE.Mess[i] ;
           END ;
        END ;
  END ;
G.TitleCenter:=TRUE ;
END ;

procedure TFCImpFic.FormShow(Sender: TObject);
Var i : Integer ;
begin
RechCompteLance:=FALSE ; Caption:=Caption+' '+StFichier ;
ModifFaite:=FALSE ;
VerifCpt.Enabled:=V_PGI.OkOuvert ;
Fillchar(GenLu,SizeOf(GenLu),#0) ; Fillchar(AuxLu,SizeOf(AuxLu),#0) ;
Fillchar(SecLu,SizeOf(SecLu),#0) ;
For i:=0 To 2 Do GP.ColAligns[i]:=taCenter ;
For i:=3 To 5 Do GP.ColAligns[i]:=taRightJustify ;
InitGrid ;
If ChargeGrid<=0 Then BEGIN VisuLignesFausses.Enabled:=FALSE ; END ;
// ajout me 13-12-01
if ctxPCL in V_PGI.PGIContexte then
   CalcNum.enabled := FALSE;
end;

procedure TFCImpFic.BValiderClick(Sender: TObject);
begin
If ModifFaite Then
   BEGIN
   END ;
Patience.Visible:=TRUE ; Application.ProcessMessages ;
Close ;
end;

procedure TFCImpFic.FormClose(Sender: TObject; var Action: TCloseAction);
begin
(*G.VidePile(False) ;*) ListeCptFaux.Clear ; ListeCptFaux.Free ;
end;

procedure TFCImpFic.BImprimerClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
PrintDBGrid(G,Nil,Caption,'') ;
{$ENDIF}
end;

procedure TFCImpFic.BRechercheClick(Sender: TObject);
begin
FindFirst:=True ; FindMvt.Execute ;
end;

procedure TFCImpFic.FindMvtFind(Sender: TObject);
begin
Rechercher(G,FindMvt,FindFirst) ;
end;

Function TFCImpFic.QuelMessSup(What : Integer ; CptLu : TCptLu) : String ;
Var Racine : String ;
BEGIN
Result:='' ;
Case What Of
  0 : BEGIN
      Racine:=Copy(CptLu.Cpt,1,1) ;
      If Racine='6' Then  Result:=' '+HMess.Mess[3] Else
         If Racine='7' Then  Result:=' '+HMess.Mess[4] ;
      If CptLu.Ventilable Then Result:=Result+' '+HMess.Mess[5] ;
      END ;
  1 : BEGIN
      Racine:=Copy(CptLu.Collectif,1,2) ;
      If Racine='40' Then Result:=' '+HMess.Mess[7] Else
         If Racine='41' Then Result:=' '+HMess.Mess[6] ;
      If CptLu.Collectif<>'' Then Result:=Result+' - '+CptLu.Collectif ;
      END ;
  END ;
END ;

{$IFNDEF EAGLCLIENT}

Function TFCImpFic.VerifCompte(What : Integer ; Var TabCpt : TTabCptLu) : Boolean ;
Var Q : Tquery ;
    i : Integer ;
    St,Cpt,MessSup : String ;
{$IFNDEF PROGEXT}
    X : DelInfo ;
{$ENDIF}
    fb : TFichierBase ;
BEGIN
{$IFNDEF PROGEXT}
Result:=FALSE ; fb:=fbgene ;
Case What Of
  0 : BEGIN
      St:='Select G_GENERAL FROM GENERAUX WHERE G_GENERAL=:CPT' ;
      fb:=fbGene ;
      END ;
  1 : BEGIN
      St:='Select T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE=:CPT' ;
      fb:=fbAux ;
      END ;
  2 : BEGIN
      St:='Select S_SECTION FROM SECTION WHERE S_SECTION=:CPT AND S_AXE="A1"' ;
      fb:=fbAxe1 ;
      END ;
  END ;
Q:=prepareSQL(St,TRUE) ;
InitMove(G.RowCount-1,'') ;
for i:=1 to MaxCptLu do
    BEGIN
    MoveCur(FALSE) ; If TabCpt[i].Cpt='' Then Break ;
    Cpt:=TabCpt[i].Cpt ; MessSup:='' ;
    If Bourre Then Cpt:=BourreOuTronque(Cpt,fb) ;
    Q.Close ;
    Q.Params[0].AsString:=Cpt ;
    Q.Open ;
    if Q.Eof then
       BEGIN
       X:=DelInfo.Create ; X.Gen:='' ; X.aux:='' ; X.Sect:='' ;
       X.GenACreer:=FALSE ; X.AuxACreer:=FALSE ; X.SectACreer:=FALSE ;
       Case What Of
          0 : BEGIN
              X.Gen:=Cpt ; X.GenACreer:=TRUE ;
              END ;
          1 : BEGIN
              X.Aux:=Cpt ; X.AuxACreer:=TRUE ;
              END ;
          2 : BEGIN
              X.Sect:=Cpt ; X.SectACreer:=TRUE ;
              END ;
          END ;
       MessSup:=QuelMessSup(What,TabCpt[i]) ;
       X.LeCod:=Cpt ; X.LeLib:=HMess.Mess[What]+MessSup ;
       Result:=TRUE ; ListeCptFaux.Add(X);
       END ;
    END ;
Ferme(Q) ;
RechCompteLance:=TRUE ;
FiniMove ;
{$ENDIF}
END ;

{$ELSE}

Function TFCImpFic.VerifCompte(What : Integer ; Var TabCpt : TTabCptLu) : Boolean ;
Var
    Q       : Tquery ;
    St,Cpt,MessSup : String ;
{$IFNDEF PROGEXT}
    X       : DelInfo ;
    i       : integer;
{$ENDIF}
    fb : TFichierBase ;
BEGIN
{$IFNDEF PROGEXT}
Result:=FALSE ; fb:=fbgene ;
Case What Of
  0 : BEGIN
      St:='Select G_GENERAL FROM GENERAUX WHERE G_GENERAL="' ;
      fb:=fbGene ;
      END ;
  1 : BEGIN
      St:='Select T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="' ;
      fb:=fbAux ;
      END ;
  2 : BEGIN
      St:='Select S_SECTION FROM SECTION WHERE  S_AXE="A1" AND S_SECTION="' ;
      fb:=fbAxe1 ;
      END ;
  END ;
InitMove(G.RowCount-1,'') ;
for i:=1 to MaxCptLu do
BEGIN
    MoveCur(FALSE) ; If TabCpt[i].Cpt='' Then Break ;
    Cpt:=TabCpt[i].Cpt ; MessSup:='' ;
    If Bourre Then Cpt:=BourreOuTronque(Cpt,fb) ;
    Q := OpenSQL(St+Cpt+'"',TRUE) ;

    if Q.Eof then
    BEGIN
       X:=DelInfo.Create ; X.Gen:='' ; X.aux:='' ; X.Sect:='' ;
       X.GenACreer:=FALSE ; X.AuxACreer:=FALSE ; X.SectACreer:=FALSE ;
       Case What Of
          0 : BEGIN
              X.Gen:=Cpt ; X.GenACreer:=TRUE ;
              END ;
          1 : BEGIN
              X.Aux:=Cpt ; X.AuxACreer:=TRUE ;
              END ;
          2 : BEGIN
              X.Sect:=Cpt ; X.SectACreer:=TRUE ;
              END ;
          END ;
       MessSup:=QuelMessSup(What,TabCpt[i]) ;
       X.LeCod:=Cpt ; X.LeLib:=HMess.Mess[What]+MessSup ;
       Result:=TRUE ; ListeCptFaux.Add(X);
    END ;
    Ferme(Q) ;
END ;
RechCompteLance:=TRUE ;
FiniMove ;
{$ENDIF}
END ;

{$ENDIF}

procedure TFCImpFic.VerifCptClick(Sender: TObject);
Var GGRien,UneErreur : Boolean ;
begin
{$IFNDEF PROGEXT}
ListeCptFaux.Clear ;
Patience.Visible:=TRUE ; Application.ProcessMessages ;
UneErreur:=VerifCompte(0,GenLu) ;
UneErreur:=VerifCompte(1,AuxLu) Or UneErreur ;
UneErreur:=VerifCompte(2,SecLu) Or UneErreur;
GGRien:=false ;
Patience.Visible:=FALSE ; Application.ProcessMessages ;
If UneErreur Then
   BEGIN
   VisuCpt.Enabled:=TRUE ;
   RapportdErreurMvt(ListeCptFaux,8,GGRien,False) ;
   END ;
{$ENDIF}
end;

Function ACopier(NumP,StNumP : String) : Boolean ;
Var i : Integer ;
BEGIN
Result:=FALSE ;
i:=Pos(NumP,StNumP) ;
If I>0 Then Result:=TRUE ;
END ;

Procedure TFCImpFic.RecopieFic(StNumP : String) ;
Var Fichier,NewFichier1,NewFichier2 : TextFile ;
    St,StNewFichier1,StNewFichier2 : String ;
    Enr : TStLue ;
    Pb : Boolean ;
BEGIN
Sauve.Title:=HMess.Mess[16] ;
if Sauve.Execute then StNewFichier1:=Sauve.FileName ;
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
AssignFile(Fichier,StFichier) ; Reset(Fichier) ;
StNewFichier2:=FileTemp('.PNM') ;
AssignFile(NewFichier1,StNewFichier1) ; Rewrite(NewFichier1) ;
AssignFile(NewFichier2,StNewFichier2) ; Rewrite(NewFichier2) ;
ReadLn(Fichier,St) ; WriteLn(NewFichier1,St) ; WriteLn(NewFichier2,St) ;
While Not EOF(Fichier) do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ; AlimEnr(St,Enr,FormatFic) ;
  Pb:=(Not EstUneLigneCpt(St)) And ACopier(Trim(Enr.NumP),StNumP) ;
  If Pb Then WriteLn(NewFichier1,St) Else WriteLn(NewFichier2,St) ;
  END ;
FiniMove ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ; CloseFile(NewFichier2) ;
AssignFile(Fichier,StFichier) ; Erase(Fichier) ;
RenameFile(StNewFichier2,StFichier) ;
END ;

procedure TFCImpFic.FormCreate(Sender: TObject);
begin
ListeCptFaux:=TList.Create ;
end;

procedure TFCImpFic.GDblClick(Sender: TObject);
Var Enr : TStLue ;
begin
If G.Row=1 Then Exit ;
RecupEnr(Enr,TRUE,G.Row) ;
If VoirEnr(Enr,taConsult) Then BEGIN AlimGrid(Enr,G.Row,FormatFic) ; ModifFaite:=TRUE ; END ;
end;

{$IFDEF EAGLCLIENT}
Procedure ChargeJalSoc(Quoi : Integer ; Var JalSoc : TTabCptLu) ;
begin
end;
{$ELSE}
Procedure ChargeJalSoc(Quoi : Integer ; Var JalSoc : TTabCptLu) ;
Var Q : TQuery ;
    i : Integer ;
    St : String ;
BEGIN
Fillchar(JalSoc,SizeOf(JalSoc),#0) ; i:=0 ;
Case Quoi Of
  0 : St:='Select * From JOURNAL' ;
  1 : St:='Select * From GENERAUX WHERE G_COLLECTIF="X"' ;
  END ;
Q:=OpenSQL(St,TRUE) ;
While Not Q.Eof Do
  BEGIN
  Inc(i) ;
  Case quoi Of
    0 : BEGIN
        JalSoc[i].Cpt:=Trim(Q.FindField('J_JOURNAL').AsString) ;
        JalSoc[i].Nature:=Q.FindField('J_NATUREJAL').AsString ;
        END ;
    1 : BEGIN
        JalSoc[i].Cpt:=Trim(Q.FindField('G_GENERAL').AsString) ;
        JalSoc[i].Nature:=Q.FindField('G_NATUREGENE').AsString ;
        END ;
    END ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;
{$ENDIF}

Function TrouveJalSoc(Cpt : String ; Var TabCpt : TTabCptLu) : TCptLu ;
Var i : Integer ;
    LeCptLu : TCptLu ;
BEGIN
Fillchar(LeCptLu,SizeOf(LeCptLu),#0) ; Result:=LeCptLu ;
For i:=1 To MaxCptLu Do
  BEGIN
  If (Cpt=TabCpt[i].Cpt) Or (TabCpt[i].Cpt='') Then BEGIN Result:=TabCpt[i] ; Exit ; END ;
  END ;
END ;


Function NbCptLu(Var TabCpt : TTabCptLu) : Integer ;
Var i : Integer ;
BEGIN
For i:=1 To MaxCptLu Do If TabCpt[i].Cpt='' Then Break ; Result:=i-1 ;
END ;

procedure TFCImpFic.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFCImpFic.TranfertFicClick(Sender: TObject);
begin
If HMess.Execute(13,'','')=mrYes Then RecopieFic(StNumP) ;
end;

procedure TFCImpFic.VisuCptClick(Sender: TObject);
Var GGRien : Boolean ;
begin
GGRien:=FALSE ;
{$IFNDEF PROGEXT}
RapportdErreurMvt(ListeCptFaux,8,GGRien,False) ;
{$ENDIF}
end;

procedure TFCImpFic.VisibleGrid (Quoi : Byte);
begin
  G.Visible:=((Quoi=0) or (Quoi=3));
  GP.Visible:=(Quoi=1) ;
  GPL.Visible:=(Quoi=2);
  GAux.Visible:=(Quoi=4) ;
  GGen.Visible:=(Quoi=5) ;
  GSect.Visible:=(Quoi=6) ;
{AJOUT ME 09-10-01}
  G7.Visible := (Quoi=7);
  G8.Visible := (Quoi=8);
  G9.Visible := (Quoi=9);
  G10.Visible := (Quoi=10);
  G11.Visible := (Quoi=11);
  G12.Visible := (Quoi=12);
end;

Procedure TFCImpFic.QuelGrid(Quoi : Byte) ;
BEGIN
{AJOUT ME 09-10-01}
VisibleGrid (Quoi);
NbPiecesFausses.Visible:=(Quoi=1) ;
NbPiecesIncorrectes.Visible:=(Quoi=1) ;
NbPieces.Visible:=(Quoi=1) ;
NbLignesFausses.Visible:=(Quoi=0) Or (Quoi=3) ;
NbLignes.Visible:=(Quoi=0) Or (Quoi=3) ;
VisuLignesFausses.Enabled:=(Quoi=0) ;
VisuBlancs.Enabled:=(Quoi=0) Or (Quoi=2) Or (Quoi=3) ;
CalcNum.Enabled:=(Quoi=0) ;

END ;

procedure TFCImpFic.VisuTotauxClick(Sender: TObject);
begin
QuelGrid(1) ;
end;

procedure TFCImpFic.VisuLignesClick(Sender: TObject);
begin
If VisuLignesFausses.Enabled=FALSE Then MasqueLigne(1) ;
QuelGrid(0) ;
end;

(*
Function TFCImpFic.FindFirstLig(StNumP : String ; Avance : Boolean) : Integer ;
BEGIN
Result:=0 ;
If Avance Then ;

END ;

procedure TFCImpFic.AlimGPL(Lig : Integer) ;
Var StNumP : String ;
    LigDep,LigFin : Integer ;
BEGIN
StNumP:=G.Cells[14,Lig] ;
LigDep:=FindFirstLig(StNumP,FALSE) ;
If LigDep=0 Then LigDep:=Lig ;
LigFin:=FindFirstLig(StNumP,TRUE) ;
If LigFin=0 Then LigFin:=Lig ;
END ;
*)

procedure TFCImpFic.VisulignespieceClick(Sender: TObject);
begin
QuelGrid(2) ; // AlimGPL(G.Row) ;
end;

Procedure TFCImpFic.MasqueLigne(Quoi : Integer) ;
Var i : Integer ;
BEGIN
If Quoi=3 Then VisuLignesFausses.Enabled:=FALSE ;
Patience.Visible:=TRUE ; Application.ProcessMessages ;
Initmove((G.RowCount-1),'') ;
For i:=1 To G.RowCount-1 Do
  BEGIN
  MoveCur(FALSE) ;
  If Quoi=3 Then
     BEGIN
     If pos('*',G.Cells[0,i])=0 Then G.RowHeights[i]:=0 Else G.RowHeights[i]:=G.DefaultRowHeight ;
     END Else
     BEGIN
     G.RowHeights[i]:=G.DefaultRowHeight ;
     END ;
  END ;
FiniMove ; Patience.Visible:=FALSE ; Application.ProcessMessages ;
END ;

procedure TFCImpFic.VisuLignesFaussesClick(Sender: TObject);
begin
QuelGrid(3) ; MasqueLigne(3) ;
end;

procedure TFCImpFic.VoirOuNonLesBlancs;
Var i,j : Integer ;
    NbCol : Integer ;
begin
Patience.Visible:=TRUE ; Application.ProcessMessages ;
NbCol:=14 ; If FormatFic=2 Then NbCol:=G.ColCount-1 ;
Initmove(NbCol*(G.RowCount-1),'') ;
For i:=1 To NbCol Do
  For j:=1 To G.RowCount-1 Do
  BEGIN
  MoveCur(FALSE) ; G.Cells[i,j]:=VoirBlancs(G.Cells[i,j],VoirBlanc.Checked) ;
  END ;
FiniMove ;
Patience.Visible:=FALSE ; Application.ProcessMessages ;
end;

procedure TFCImpFic.VisuBlancsClick(Sender: TObject);

begin
VoirBlanc.Checked:=Not VoirBlanc.Checked ; VoirOuNonLesBlancs ;
If VoirBlanc.Checked Then VisuBlancs.Hint:=HMess.Mess[9] Else VisuBlancs.Hint:=HMess.Mess[8] ;
end;

procedure TFCImpFic.CalcNumClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
Var
    StErr : String ;
    OkOk  : Boolean ;
    InfoImp : TInfoImport ;
    QFiche : TQFiche ;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
FillChar(InfoImp,SizeOf(InfoImp),#0) ; InfoImp.NomFic:=StFichier ; StErr:='' ;
FillChar(QFiche,SizeOf(QFiche),#0) ;
If FormatFic=0 Then InfoImp.Format:='SAA' Else
 If FormatFic=1 Then InfoImp.Format:='HLI' Else
  If FormatFic=2 Then InfoImp.Format:='HAL' Else
   If FormatFic=3 Then InfoImp.Format:='CGN' Else
    If FormatFic=4 Then InfoImp.Format:='CGE' Else Exit ;
If HMess.Execute(11,'','')=mrYes Then
  BEGIN
  Case HMess.Execute(12,'','') Of mrYes, mrCancel : Exit ; END ;
  Patience.Visible:=TRUE ; Application.ProcessMessages ;
  OkOk:=TravailleFichier (TRUE,StErr,InfoImp,QFiche,Ignorel1) ;
  If OkOk Then HMess.Execute(14,'','') ;
  Patience.Visible:=FALSE ; Application.ProcessMessages ;
  END ;
{$ENDIF}
end;

function fileTemp1(const aExt: String): String;
var
  Buffer: array[0..1023] of Char;
  aFile : String;
begin
  GetTempPath(Sizeof(Buffer)-1,Buffer);
  GetTempFileName(Buffer,'TMP',1,Buffer);
  SetString(aFile, Buffer, StrLen(Buffer));
  Result:=ChangeFileExt(aFile,aExt);
end;

{$IFNDEF EAGLCLIENT}
Procedure FicToFicCorresp(Quoi,FormatFic : Integer ; StFichier,StNewFichier1 : String) ;
Var Fichier,NewFichier1 : TextFile ;
    St,StNewFichier2,SQL,CRTYPE,Cpt : String ;
    Enr : TStLue ;
    Q : TQuery ;
BEGIN
AssignFile(Fichier,StFichier) ;
{$I-} Reset(Fichier) ; {$I+} if IoResult<>0 then Exit ;
AssignFile(NewFichier1,StNewFichier1) ; Rewrite(NewFichier1) ;
InitMove(1000,'') ;
ReadLn(Fichier,St) ; WriteLn(NewFichier1,St) ;
Case Quoi of
  2 : CRTYPE:='IA1' ;
  END ;
SQL:='SELECT CR_TYPE,CR_CORRESP,CR_LIBELLE from CORRESP WHERE CR_TYPE="'+CRTYPE+'" AND CR_CORRESP=:CPT' ;
Q:=PrepareSQL(SQL,TRUE) ;
While Not EOF(Fichier) do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  If (Trim(St)<>'') And Not(EstUneLigneCpt(St)) Then
     BEGIN
     AlimEnr(St,Enr,FormatFic) ;
     If Enr.TC='A' Then
        BEGIN
        Q.Close ;
        Q.Params[0].AsString:=Trim(Enr.AuxSect) ;
        Q.Open ;
        If Not Q.Eof Then
           BEGIN
           Cpt:=Format_String(Q.Fields[2].AsString,13) ;
           Case FormatFic Of
             0 : BEGIN
                 Delete(St,26,13) ;
                 Insert(Cpt,St,26) ;
                 END ;
             END ;
           END ;
        END ;
     WriteLn(NewFichier1,St) ;
     END ;
  END ;
FiniMove ;
Q.Close ; Ferme(Q) ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ;
StNewFichier2:=FileTemp1('.PNM') ;
RenameFile(StFichier,StNewFichier2) ;
RenameFile(StNewFichier1,StFichier) ;
RenameFile(StNewFichier2,StNewFichier1) ;
END ;
Procedure PositionneSurAttente(Bourre : Boolean ; Quoi : Integer ; FormatFic : Integer ; StFichier : String) ;
Var Fichier,NewFichier : TextFile ;
    NewFileName : String ;
    St,Gen,Aux,TC,AuxRemplace,NomC,NomT : String ;
    DebGen,LgGen,DebAux,LgAux,DebTC : Integer ;
    Enr : TStLue ;
    CptColl : TTabCptLu ;
    LeCptColl : TCptLu ;
    OkColl : Boolean ;
    Q : TQuery ;
    SQL : String ;
BEGIN
DebTC := 0; LgAux := 0; DebAux := 0; DebGen := 0; LgGen := 0;
AssignFile(Fichier,StFichier) ; Reset(Fichier) ;
NewFileName:=FileTemp('.PNM') ;
AssignFile(NewFichier,NewFileName) ; Rewrite(NewFichier) ;
ChargeJalSoc(1,CptColl) ;
ReadLn(Fichier,St) ; WriteLn(NewFichier,St) ;
Case FormatFic Of
  0 : BEGIN { SAARI COEUR DE GAMME  mais pas NEGOCE V2 }
      DebGen:=12 ; LgGen:=13 ; DebAux:=26 ; LgAux:=13 ; DebTC:=25 ;
      END ;
  1 : BEGIN { HALLEY LIGHT }
      DebGen:=14 ; LgGen:=17 ; DebAux:=32 ; LgAux:=17 ; DebTC:=31 ;
      END ;
  2 : BEGIN { HALLEY ETENDU }
      DebAux:=14 ; LgAux:=17 ; DebTC:=31 ; LgAux:=17 ; DebTC:=31 ;
      END ;
  3 : BEGIN { CEGID NORMAL }
      END ;
  END ;
Case Quoi Of
  1 : BEGIN NomT:='TIERS' ; NomC:='T_AUXILIAIRE' ; END ;
  END ;
SQL:='SELECT '+NomC+' FROM '+NomT+' Where '+NomC+'=:CPT ' ;
Q:=PrepareSQL(SQL,TRUE) ;
InitMove(1000,'') ;
While Not EOF(Fichier) do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ; AlimEnr(St,Enr,FormatFic) ;
  If Trim(St)<>'' Then
     BEGIN
     If Not EstUneLigneCpt(St) Then
       BEGIN
       Gen:=Trim(Copy(St,DebGen,LgGen)) ;
       Aux:=Trim(Copy(St,DebAux,LgAux)) ;
       TC:=Trim(Copy(St,DebTC,1)) ;
       If Bourre Then Gen:=BourreOuTronque(Gen,fbGene) ;
       If Bourre Then Aux:=BourreOuTronque(Aux,fbAux) ;
       LeCptColl:=TrouveJalSoc(Gen,CptColl) ;
       If LeCptColl.Cpt<>'' Then
          BEGIN
          OkColl:=(LeCptColl.Nature='COC') Or (LeCptColl.Nature='COF') Or (LeCptColl.Nature='COS') Or (LeCptColl.Nature='COD') ;
          If OkColl And (Enr.TC='X') Then
             BEGIN
             Q.Close ;
             Q.Params[0].AsString:=Aux ;
             Q.Open ;
             If Q.Eof Then Aux:='' ;
             END ;
          If (Aux='') And OkColl Then
             BEGIN
             If LeCptColl.Nature='COC' Then AuxRemplace:='CLIATTENT    ' Else
             If LeCptColl.Nature='COF' Then AuxRemplace:='FOUATTENT    ' Else
             If LeCptColl.Nature='COS' Then AuxRemplace:='SALATTENT    ' ;
             Case FormatFic Of
               0 : BEGIN
                   Delete(St,DebAux,LgAux) ;
                   Insert(AuxRemplace,St,DebAux) ;
                   END ;
               End ;
             Delete(St,DebTC,1) ;
             Insert('X',St,DebTC) ;
             END ;
          END ;
       END ;
     WriteLn(NewFichier,St) ;
     END ;
  END ;
FiniMove ;
Q.Close ; Ferme(Q) ;
CloseFile(Fichier) ; CloseFile(NewFichier) ;
AssignFile(Fichier,StFichier) ; Erase(Fichier) ;
RenameFile(NewFileName,StFichier) ;
END ;

{$ENDIF}

procedure TFCImpFic.CorrespClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
Var StNewFichier1 : String;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
If HMess.Execute(15,'','')=mrYes Then
   BEGIN
   Sauve.Title:=HMess.Mess[17] ;
   if Sauve.Execute then StNewFichier1:=Sauve.FileName Else Exit ;
   Patience.Visible:=TRUE ; Application.ProcessMessages ;
   FicToFicCorresp(2,FormatFic,StFichier,StNewFichier1) ;
   Patience.Visible:=FALSE ; Application.ProcessMessages ;
   END ;
{$ENDIF}
end;

Procedure RecalculTC(FormatFic : Integer ; StFichier : String) ;
Var Fichier,NewFichier : TextFile ;
    NewFileName : String ;
    St,Gen,Aux,TC : String ;
    DebGen,LgGen,DebAux,LgAux,DebTC : Integer ;
    Enr : TStLue ;
BEGIN
DebGen := 0;  LgGen := 0; DebTC:=0 ; LgAux:=0 ;  DebAux:=0 ;
AssignFile(Fichier,StFichier) ; Reset(Fichier) ;
NewFileName:=FileTemp('.PNM') ;
AssignFile(NewFichier,NewFileName) ; Rewrite(NewFichier) ;
ReadLn(Fichier,St) ; WriteLn(NewFichier,St) ;
Case FormatFic Of
  0 : BEGIN { SAARI COEUR DE GAMME  mais pas NEGOCE V2 }
      DebGen:=12 ; LgGen:=13 ; DebAux:=26 ; LgAux:=13 ; DebTC:=25 ;
      END ;
  1 : BEGIN { HALLEY LIGHT }
      DebGen:=14 ; LgGen:=17 ; DebAux:=32 ; LgAux:=17 ; DebTC:=31 ;
      END ;
  2 : BEGIN { HALLEY ETENDU }
      DebAux:=14 ; LgAux:=17 ; DebTC:=31 ; LgAux:=17 ; DebTC:=31 ;
      END ;
  Else Exit ;
  END ;
While Not EOF(Fichier) do
  BEGIN
  ReadLn(Fichier,St) ; AlimEnr(St,Enr,FormatFic) ;
  If Trim(St)<>'' Then
     BEGIN
     If Not EstUneLigneCpt(St) Then
       BEGIN
       Gen:=Trim(Copy(St,DebGen,LgGen)) ;
       Aux:=Trim(Copy(St,DebAux,LgAux)) ;
       TC:=Trim(Copy(St,DebTC,1)) ;
       If (Aux='') And (TC<>'') then
         BEGIN
         Delete(St,DebTC,1) ;
         Insert(' ',St,DebTC) ;
         END ;
       If (Aux<>'') And (TC='') then
         BEGIN
         Delete(St,DebTC,1) ;
         Insert('X',St,DebTC) ;
         END ;
       END ;
     WriteLn(NewFichier,St) ;
     END ;
  END ;
CloseFile(Fichier) ; CloseFile(NewFichier) ;
AssignFile(Fichier,StFichier) ; Erase(Fichier) ;
RenameFile(NewFileName,StFichier) ;
END ;




procedure TFCImpFic.CalcTCClick(Sender: TObject);
begin
If HMess.Execute(18,'','')=mrYes Then
   BEGIN
   Patience.Visible:=TRUE ; Application.ProcessMessages ;
   RecalculTC(FormatFic,StFichier) ;
   Patience.Visible:=FALSE ; Application.ProcessMessages ;
   HMess.Execute(14,'','') ;
   END ;
end;

procedure TFCImpFic.CalcAttentClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
If HMess.Execute(19,'','')=mrYes Then
   BEGIN
   Patience.Visible:=TRUE ; Application.ProcessMessages ;
   PositionneSurAttente(Bourre,1,FormatFic,StFichier) ;
   Patience.Visible:=FALSE ; Application.ProcessMessages ;
   HMess.Execute(14,'','') ;
   END ;
{$ENDIF}
end;

procedure TFCImpFic.VisuAuxClick(Sender: TObject);
begin
QuelGrid(4) ;
end;

procedure TFCImpFic.VisuGenClick(Sender: TObject);
begin
QuelGrid(5) ;
end;

{AJOUT ME 09-10-01}
procedure TFCImpFic.VisuGerer1Click(Sender: TObject);
begin
QuelGrid(6) ;
end;

procedure TFCImpFic.VisuGerer7Click(Sender: TObject);
begin
QuelGrid(7) ;
end;
{fin AJOUT ME 09-10-01}

procedure TFCImpFic.GRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
Var StNum : String ;
begin
WhatError.Caption:='' ;
If pos('*',G.Cells[0,Ou])<>0 Then
  BEGIN
  StNum:='N° '+Trim(G.Cells[14,Ou]) ;
  If pos('J',G.Cells[0,Ou])<>0 Then WhatError.Caption:='Erreur : Pièce '+StNum+' avec des journaux différents' ;
  If pos('D',G.Cells[0,Ou])<>0 Then WhatError.Caption:='Erreur : Pièce '+StNum+' avec des dates différentes' ;
  If pos('P',G.Cells[0,Ou])<>0 Then WhatError.Caption:='Erreur : Pièce '+StNum+' avec des types de pièce différents' ;
  If pos('T',G.Cells[0,Ou])<>0 Then WhatError.Caption:='Erreur : Pièce '+StNum+' avec des types d''écritures différents' ;
  If pos('$',G.Cells[0,Ou])<>0 Then WhatError.Caption:='Erreur : Pièce '+StNum+' avec des devises différentes' ;
  If pos('X',G.Cells[0,Ou])<>0 Then WhatError.Caption:='Erreur : Pièce '+StNum+' avec des taux différents' ;
  If pos('E',G.Cells[0,Ou])<>0 Then WhatError.Caption:='Erreur : Pièce '+StNum+' avec des établissements différents' ;
  END ;
end;
{AJOUT ME 09-10-01}
procedure TFCImpFic.VisuGener2Click(Sender: TObject);
begin
QuelGrid(8) ;
end;

procedure TFCImpFic.VisuGener4Click(Sender: TObject);
begin
QuelGrid(9) ;
end;

procedure TFCImpFic.VisuGener5Click(Sender: TObject);
begin
QuelGrid(10) ;
end;

procedure TFCImpFic.VisuGener6Click(Sender: TObject);
begin
QuelGrid(11) ;
end;

procedure TFCImpFic.VisuExercClick(Sender: TObject);
begin
QuelGrid(12) ;
end;
{FIN AJOUT ME 09-10-01}

end.


