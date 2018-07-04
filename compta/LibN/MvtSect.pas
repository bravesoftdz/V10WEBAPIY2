unit MvtSect;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mask, Hctrls, StdCtrls, Hcompte, HTB97, ComCtrls, Ent1, HEnt1,
  hmsgbox, HSysMenu, ParamDat, HStatus, ExtCtrls, CritEdt, Ed_Tools,
  Spin, Menus, UObjFiltres {YMO 15/05/2006 FQ18127 Nouvelle gestion des filtres}, TImpFic, ImpFicU, UTOB, db, uLibWindows
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
  , HDB
  {$ENDIF}
  ;

Type tCritCeric = Record
                  Lg1,Lg2 : Integer ;
                  Cpt1,Cpt2,Axe : String ;
                  TLS : ARRAY[0..9] Of String ;
                  TLG : ARRAY[0..9] Of String ;
                  Joker : Boolean ;
                  DateOuvre,DateFerme : TDateTime ;
                  VoirDateOuvre,VoirDateFerme : Boolean ;
                  Sauf : String ;
                  End ;

Type tExtractCeric = Record
                     Lg1,Lg2 : Integer ;
                     Cpt1,Cpt2,Axe : String ;
                     TLS : ARRAY[0..9] Of String ;
                     TLG : ARRAY[0..9] Of String ;
                     Joker : Boolean ;
                     Exo : String ;
                     DateE1,DateE2 : tDateTime ;
                     Sauf : String ;
                     OkSim,OkSit,OkRev,OkPrev,OkIAIF, OkNormale : Boolean ;
                     End ;

Type tGenereCERIC = Record
                    JalG,QualG,NomFic,CptCB,CptCG,CptPB,CptPG : String ;
                    ExoG : String ;
                    SoucheN,SoucheS : String ;
                    DateG : TDateTime ;
                    NOMCHAMPTABLE : String ;
                    GTABLEC,GTABLEP : String ;
                    Lib : String ;
                    SectProv : String ;
                    SQL : String ;
                    End ;

Type tCritGenereCERIC = Record
                        E : tExtractCeric ;
                        G : tGenereCERIC ;
                        End ;

type tcritexpmvt = Record
                   Format,Exo,Etab,Jal1,Jal2,TypeEcr : String ;
                   Date1,Date2 : tdatetime ;
                   NumP1,NumP2 : Integer ;
                   NomFic : String ;
                   OkExport : Boolean ;
                   AvecODA : Boolean ;
                   End ;


Type TTL = Array[0..9] Of String ;
Procedure GenereLesEcritures(Var LeCrit : TCritCeric) ;
procedure GeneMvt;

type
  TFGenere = class(TForm)
    Pages: TPageControl;
    Extraction: TTabSheet;
    Dock971: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    PanelFiltre: TToolWindow97;
    HS_SECTIONJOKER: THLabel;
    HS_SECTION: THLabel;
    HS_SECTION_: TLabel;
    TS_AXE: THLabel;
    S_AXE: THValComboBox;
    HLabel1: THLabel;
    FExcep: TEdit;
    HLabel4: THLabel;
    FExercice: THValComboBox;
    HLabel6: THLabel;
    FDateExtract1: TMaskEdit;
    Label7: TLabel;
    FDateExtract2: TMaskEdit;
    Generation: TTabSheet;
    HLabel2: THLabel;
    FDateGenere: TMaskEdit;
    TFGen: THLabel;
    FJalGenere: THCpteEdit;
    TSelectCpte: THLabel;
    FQualGenere: THValComboBox;
    LFile: THLabel;
    FileName: TEdit;
    RechFile: TToolbarButton97;
    Sauve: TSaveDialog;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    GroupBox1: TGroupBox;
    HLabel3: THLabel;
    FCptCB: THCpteEdit;
    HLabel5: THLabel;
    FCptCG: THCpteEdit;
    GroupBox2: TGroupBox;
    HLabel7: THLabel;
    HLabel8: THLabel;
    FCptPB: THCpteEdit;
    FCptPG: THCpteEdit;
    TG_TABLE11: THLabel;
    HLabel9: THLabel;
    G_TABLEP: THCpteEdit;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    LEnCours: TLabel;
    Pzlibre: TTabSheet;
    TS_TABLE0: TLabel;
    S_TABLE0: THCpteEdit;
    TS_TABLE1: TLabel;
    S_TABLE1: THCpteEdit;
    TS_TABLE2: TLabel;
    S_TABLE2: THCpteEdit;
    TS_TABLE3: TLabel;
    S_TABLE3: THCpteEdit;
    TS_TABLE4: TLabel;
    S_TABLE4: THCpteEdit;
    TS_TABLE5: TLabel;
    S_TABLE5: THCpteEdit;
    TS_TABLE6: TLabel;
    S_TABLE6: THCpteEdit;
    TS_TABLE7: TLabel;
    S_TABLE7: THCpteEdit;
    TS_TABLE8: TLabel;
    S_TABLE8: THCpteEdit;
    TS_TABLE9: TLabel;
    S_TABLE9: THCpteEdit;
    G_TABLEC: THCpteEdit;
    GenTlibre: TTabSheet;
    G_TABLE0: THCpteEdit;
    TG_TABLE0: TLabel;
    TG_TABLE1: TLabel;
    G_TABLE1: THCpteEdit;
    TG_TABLE2: TLabel;
    G_TABLE2: THCpteEdit;
    G_TABLE3: THCpteEdit;
    TG_TABLE3: TLabel;
    TG_TABLE4: TLabel;
    G_TABLE4: THCpteEdit;
    TG_TABLE5: TLabel;
    G_TABLE5: THCpteEdit;
    G_TABLE6: THCpteEdit;
    TG_TABLE6: TLabel;
    G_TABLE7: THCpteEdit;
    TG_TABLE7: TLabel;
    G_TABLE8: THCpteEdit;
    TG_TABLE8: TLabel;
    G_TABLE9: THCpteEdit;
    TG_TABLE9: TLabel;
    FQualifPiece: THMultiValComboBox;
    HLabel10: THLabel;
    FFiltres: THValComboBox;
    S_SECTION: THCritMaskEdit;
    S_SECTION_: THCritMaskEdit;
    procedure FExerciceChange(Sender: TObject);
    procedure FDateExtract1KeyPress(Sender: TObject; var Key: Char);
    procedure RechFileClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure S_AXEChange(Sender: TObject);
    procedure S_SECTIONChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    {
    procedure POPFPopup(Sender: TObject);
    }
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure S_SECTIONExit(Sender: TObject);
  private
    { Déclarations privées }
    ObjFiltre : TObjFiltre;
    LeCrit : TCritCERIC ;
    CritG : tCritGenereCERIC ;
    procedure RecupCrit ;
    Function  CritOk(Var LeCrit : tCritGenereCERIC) : Boolean ;
    Function GenereMvt(L : tStringList ; Var LeCrit : tCritGenereCERIC ; Var CritExpMvt : TCritExpMvt ; Sens : String) : Integer ;

  public
    { Déclarations publiques }
  end;


implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPProcGen,
  ULibExercice,
  {$ENDIF MODENT1}
  SAISUTIL, SAISCOMM, CPTEUTIL, LetBatch, Utilsais ;

{$R *.DFM}

Procedure GenereLesEcritures(Var LeCrit : TCritCeric) ;
var XX : TFGenere ;
BEGIN
XX:=TFGenere.Create(Application) ;
 Try
  XX.LeCrit:=LeCrit ;
  XX.ShowModal ;
 Finally
  XX.Free ;
 End ;
SourisNormale ;
END ;

procedure GeneMvt;
Var LeCrit : TCritCeric ;
begin
Fillchar(LeCrit,SizeOf(LeCrit),#0) ;
GenereLesEcritures(LeCrit) ;
end;

procedure TFGenere.FExerciceChange(Sender: TObject);
begin
If FExercice.ItemIndex>0 Then ExoToDates(FExercice.Value,FDateExtract1,FDateExtract2) ;
end;

procedure TFGenere.FDateExtract1KeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TFGenere.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFGenere.FormShow(Sender: TObject);
begin

G_TABLEC.ZoomTable:=tzRubCpta ;
G_TABLEP.ZoomTable:=tzRubCpta ;

S_AXE.Value:=LeCrit.Axe ;
If S_AXE.ItemIndex=-1 Then S_AXE.Value:='A1' ;
S_SECTION.Text:=LeCrit.Cpt1 ;
S_SECTION_.Text:=LeCrit.Cpt2 ;
//NbCharRupt1.Value:=LeCrit.Lg1 ;
//NbCharRupt1.Value:=LeCrit.Lg2 ;
FExcep.Text:=LeCrit.Sauf ;
FExercice.Value:=VH^.Entree.Code ;
FQualGenere.ItemIndex:=0 ;
FDateGenere.Text:=DateToStr(FinDemois(V_PGI.DateEntree)) ;
Pages.ActivePage:=Extraction ;
UpdateCaption(Self) ;

//YMO 05/06/2006 Déplacement du chargement du filtre
//YMO 15/05/2006 FQ18127 Nouvelle gestion des filtres
ObjFiltre.Charger;

end;

procedure TFGenere.S_AXEChange(Sender: TObject);
begin
// YMO 28/04/2006 FQ17954
If Not ObjFiltre.InChargement then
begin
  S_SECTION.Clear ;
  S_SECTION_.Clear ;
//  FJoker.Clear ;
end;
Case S_AXE.ItemIndex of
(*
 0 : BEGIN S_SECTION.ZoomTable:=tzSection  ; S_SECTION_.ZoomTable:=tzSection   ; END ;
 1 : BEGIN S_SECTION.ZoomTable:=tzSection2 ; S_SECTION_.ZoomTable:=tzSection2 ; END ;
 2 : BEGIN S_SECTION.ZoomTable:=tzSection3 ; S_SECTION_.ZoomTable:=tzSection3 ; END ;
 3 : BEGIN S_SECTION.ZoomTable:=tzSection4 ; S_SECTION_.ZoomTable:=tzSection4 ; END ;
 4 : BEGIN S_SECTION.ZoomTable:=tzSection5 ; S_SECTION_.ZoomTable:=tzSection5 ; END ;
 *)
 0 : BEGIN S_SECTION.Plus:='S_AXE="A1" AND S_AFFAIREENCOURS="X"'  ; S_SECTION_.Plus:='S_AXE="A1" AND S_AFFAIREENCOURS="X"'   ; END ;
 1 : BEGIN S_SECTION.Plus:='S_AXE="A2" AND S_AFFAIREENCOURS="X"'  ; S_SECTION_.Plus:='S_AXE="A2" AND S_AFFAIREENCOURS="X"'   ; END ;
 2 : BEGIN S_SECTION.Plus:='S_AXE="A3" AND S_AFFAIREENCOURS="X"'  ; S_SECTION_.Plus:='S_AXE="A3" AND S_AFFAIREENCOURS="X"'   ; END ;
 3 : BEGIN S_SECTION.Plus:='S_AXE="A4" AND S_AFFAIREENCOURS="X"'  ; S_SECTION_.Plus:='S_AXE="A4" AND S_AFFAIREENCOURS="X"'   ; END ;
 4 : BEGIN S_SECTION.Plus:='S_AXE="A5" AND S_AFFAIREENCOURS="X"'  ; S_SECTION_.Plus:='S_AXE="A5" AND S_AFFAIREENCOURS="X"'   ; END ;
 End ;

end;

procedure TFGenere.RecupCrit ;
Var i : Integer ;
    MonoExo : Boolean ;
    Exo1 : TExoDate ;
    TH : TEdit ;
    TLL : TLabel ;
BEGIN
Fillchar(CritG,SizeOf(CritG),#0) ;
With CritG.E Do
  BEGIN
  
  Axe:=S_AXE.Value ;
  
//  Lg1:=NbCharRupt1.Value ; Lg2:=NbCharRupt2.Value ;
//  If FJoker.Visible Then
    if TestJoker(S_SECTION.Text) then
    BEGIN
    Joker:=TRUE ;
//    If Trim(FJoker.Text)<>'' Then BEGIN Cpt1:=Trim(FJoker.Text) ; Cpt2:=Trim(FJoker.Text) ; END ;
    If Trim(S_SECTION.Text)<>'' Then BEGIN Cpt1:=Trim(S_SECTION.Text) ; Cpt2:=Trim(S_SECTION.Text) ; END ;
    END Else
    BEGIN
    Joker:=FALSE ;
    If Trim(S_SECTION.Text)<>'' Then Cpt1:=Trim(S_SECTION.Text) ;
    If Trim(S_SECTION_.Text)<>'' Then Cpt2:=Trim(S_SECTION_.Text) ;
    END ;
  If FExercice.ItemIndex>0 Then Exo:=FExercice.Value ;
  DateE1:=StrToDate(FDateExtract1.Text) ;
  DateE2:=StrToDate(FDateExtract2.Text) ;
  If Trim(FExcep.Text)<>'' Then Sauf:=Trim(FExcep.Text) ;
  //YMO 28/04/2006 FQ17955 Nouveau Combo Type d'écriture
  OkNormale:=(Pos('N', UpperCase(FQualifPiece.Text)) > 0) Or (FQualifPiece.Text='<<Tous>>');  
  OkSim:=(Pos('S', UpperCase(FQualifPiece.Text)) > 0) Or (FQualifPiece.Text='<<Tous>>');
  OkSit:=(Pos('U', UpperCase(FQualifPiece.Text)) > 0) Or (FQualifPiece.Text='<<Tous>>');
  OkPrev:=(Pos('P', UpperCase(FQualifPiece.Text)) > 0) Or (FQualifPiece.Text='<<Tous>>');
  OkRev:=(Pos('R', UpperCase(FQualifPiece.Text)) > 0) Or (FQualifPiece.Text='<<Tous>>');
  OkIAIF:=(Pos('I', UpperCase(FQualifPiece.Text)) > 0) Or (FQualifPiece.Text='<<Tous>>');

    If PZLibre.TabVisible Then For i:=0 To 9 Do
    BEGIN
    TLL:=TLabel(Self.FindComponent('TS_TABLE'+IntToStr(i))) ;
    if TLL<>Nil then
       BEGIN
       TH:=TEdit(TLL.FocusControl) ; if TH<>Nil then If Trim(TH.Text)<>'' Then TLS[i]:=Trim(TH.Text)
       END ;
    END ;

    If GenTLibre.TabVisible Then For i:=0 To 9 Do
    BEGIN
    TLL:=TLabel(Self.FindComponent('TG_TABLE'+IntToStr(i))) ;
    if TLL<>Nil then
       BEGIN
       TH:=TEdit(TLL.FocusControl) ; if TH<>Nil then If Trim(TH.Text)<>'' Then TLG[i]:=Trim(TH.Text)
       END ;
    END ;

  END ;
With CritG.G Do
  BEGIN
  JalG:=FJalGenere.text ; QualG:=FQualGenere.Value ; NomFic:=FileName.text ;
  CptCB:=FCptCB.Text ; CptCG:=FCptCG.Text ; CptPB:=FCptPB.Text ; CptPG:=FCptPG.Text ;
  DateG:=StrToDate(FDateGenere.text) ;
  QuelExoDate(DateG,DateG,MonoExo,Exo1) ; ExoG:=Exo1.Code ;
  
  GTABLEC:=G_TABLEC.Text ; GTABLEP:=G_TABLEP.Text ;
  END ;
END ;

Function ExisteCpt(Cpt : String ; Var Vent : Boolean) : Boolean ;
Var Q : TQuery ;
    St: String;
BEGIN
Result:=TRUE ; Vent:=FALSE ;
St:='SELECT G_GENERAL,G_VENTILABLE1 FROM GENERAUX WHERE G_GENERAL="'+Cpt+'" ';
Q:=OpenSQL(St, TRUE) ;

If Q.Eof Then Result:=FALSE Else Vent:=Q.FindField('G_VENTILABLE1').AsString='X' ;
Ferme(Q) ;
END ;

Function TFGenere.CritOk(Var LeCrit : tCritGenereCERIC) : Boolean ;
Var NumErr : Integer ;
    Q : TQuery ;
    Vent : Boolean ;
BEGIN
Result:=FALSE ; NumErr:=0 ;
Q:=OpenSQL('SELECT J_NATUREJAL,J_COMPTEURNORMAL,J_COMPTEURSIMUL FROM JOURNAL WHERE J_JOURNAL="'+CritG.G.JalG+'" ',TRUE) ;
If Q.Eof Then NumErr:=1 Else
  BEGIN
  LeCrit.G.SoucheN:=Q.FindField('J_COMPTEURNORMAL').AsString ;
  LeCrit.G.SoucheS:=Q.FindField('J_COMPTEURSIMUL').AsString ;
  If (LeCrit.G.SoucheN='') And (LeCrit.G.QualG='N') Then NumErr:=6 Else
    If (LeCrit.G.SoucheN='') And (LeCrit.G.QualG<>'N') Then NumErr:=7 ;
  END ;
Ferme(Q) ;
If Not ExisteCpt(LeCrit.G.CptCB,Vent) Then NumErr:=2 Else If Vent Then NumErr:=15 ;
If Not ExisteCpt(LeCrit.G.CptCG,Vent) Then NumErr:=3 Else If Not Vent Then NumErr:=16 ;
If Not ExisteCpt(LeCrit.G.CptPB,Vent) Then NumErr:=4 Else If Vent Then NumErr:=17 ;
If Not ExisteCpt(LeCrit.G.CptPG,Vent) Then NumErr:=5 Else If Not Vent Then NumErr:=18 ;
//If LeCrit.G.NOMCHAMPTABLE='' Then NumErr:=8 ;
If LeCrit.G.GTABLEC='' Then NumErr:=9 ;
If LeCrit.G.GTABLEP='' Then NumErr:=10 ;
if LeCrit.G.ExoG = '' then NumErr := 21;
If NumErr>0 Then HM.Execute(NumErr,Caption,'') Else Result:=TRUE ;
{ YMO 04/2006 Pas d'export fichier
If Result And (LeCrit.G.NomFic='') Then
  BEGIN
  If HM.Execute(14,Caption,'')<>mrYes Then Result:=FALSE ;
  END ;
}
END ;

procedure TFGenere.S_SECTIONChange(Sender: TObject);
Var AvecJoker : Boolean ;
begin
//AvecJoker:=Joker(S_SECTION, S_SECTION_, FJoker) ;
  AvecJoker :=  TestJoker(S_SECTION.Text);
  HS_SECTION_.Visible:=Not AvecJoker ;
  HS_SECTION.Visible:=Not AvecJoker ;
  HS_SECTIONJOKER.Visible:=AvecJoker ;
end;

Function LeWhere(Var LeCrit : tCritGenereCERIC) : String ;
Var St : String ;
    i : Integer ;
BEGIN
St:='' ;
St:=St+' S_AXE="'+LeCrit.E.Axe+'" ' ;
If LeCrit.E.Joker Then
  BEGIN
  St:=St+'AND S_SECTION like "'+TraduitJoker(LeCrit.E.Cpt1)+'" ' ;
  END Else
  BEGIN
  If LeCrit.E.Cpt1<>'' Then St:=St+' AND S_SECTION>="'+LeCrit.E.Cpt1+'" ' ;
  If LeCrit.E.Cpt2<>'' Then St:=St+' AND S_SECTION<="'+LeCrit.E.Cpt2+'" ' ;
  END ;

For i:=0 To 9 Do
  If LeCrit.E.TLS[i]<>'' Then St:=St+' AND S_TABLE'+IntToStr(i)+'="'+LeCrit.E.TLS[i]+'" ' ;

If LeCrit.E.Sauf<>'' Then St:=St+' AND '+AnalyseCompte(LeCrit.E.Sauf,AxeToFb(LeCrit.E.Axe),True,False) ;

St:=' WHERE '+St+' ' ;
Result:=St ;
END ;

Procedure FaitListeRacine(L : tStringList ; Var LeCrit : tCritGenereCERIC) ;
Var Q : TQuery ;
    Aff,St,StWhere : String ;
    AffATraiter : Boolean ;
BEGIN
// Recense toutes les sections remplissant les conditions sélectionnées à l'écran
// qui n'ont pas de date de fin de chantier (ouvertes)
StWhere:=LeWhere(LeCrit) ;
St:='SELECT S_SECTION,S_AXE,S_DEBCHANTIER, S_FINCHANTIER FROM SECTION '+StWhere+' ORDER BY S_SECTION' ; Aff:='' ;
Q:=OpenSQL(St,TRUE) ;
InitMove(RecordsCount(Q),'') ; Aff:='' ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  St:=Q.Fields[0].AsString ; Aff:=St ;
  AffATraiter:=FALSE ;
  If (Q.FindField('S_DEBCHANTIER').AsDateTime<>IDate1900)
//  YMO FQ17947 22/05/2006 On ne prenait que les sections qui n'avaient pas de date de fin de chantier
  And (Q.FindField('S_FINCHANTIER').AsDateTime>LeCrit.G.DateG) Then AffATraiter:=TRUE ;
  If AffATraiter Then L.Add(Aff) ;
  Q.Next ;
  END ;
FiniMove ;
Ferme(Q) ;
END ;


Function QuelLib(St : String ; Var TL :TTL) : String ;
Var Q : TQuery ;
    St1 : String ;
    i : Integer ;
BEGIN
Result:='' ;
St1:='SELECT S_LIBELLE,S_TABLE0, S_TABLE1, S_TABLE2, S_TABLE3, S_TABLE4, '+
     'S_TABLE5, S_TABLE6, S_TABLE7, S_TABLE8, S_TABLE9  FROM SECTION WHERE S_SECTION="'+ST+'" ' ;
Q:=OpenSQL(St1,TRUE) ;
If Not Q.Eof Then
  BEGIN
  Result:=Q.Fields[0].AsString ;
  For i:=0 To 9 Do TL[i]:=Q.FindField('S_TABLE'+IntToStr(i)).AsString ;
  END ;
Ferme(Q) ;
END ;

Procedure VerifSectProv(Var LeCrit : tCritGenereCERIC ; Var TL : TTL) ;
Var Sect : String ;
    Q : TQuery ;
    i : Integer ;
{$IFDEF EAGLCLIENT}
    Q1 : TOB ;
{$ENDIF}
BEGIN
Sect:=LeCrit.G.SectProv ;
Q := OpenSQL('SELECT * FROM SECTION WHERE S_SECTION="'+Sect+'" ',FALSE) ;
If Q.Eof Then
  BEGIN
{$IFDEF EAGLCLIENT}
  Q1 := TOB.Create('SECTION',nil,-1) ;
  Q1.PutValue('S_SECTION', Sect) ;
  Q1.PutValue('S_AXE', LeCrit.E.Axe) ;
  Q1.PutValue('S_LIBELLE', 'Pour Solde (Situation)') ;
  Q1.PutValue('S_ABREGE', 'Situation') ;
  Q1.PutValue('S_CONFIDENTIEL', '1') ;

  For i:=0 To 9 Do Q1.PutValue('S_TABLE'+IntToStr(i), TL[i]) ;

  Q1.InsertDB(nil);
  Q1.Free;
{$ELSE}
  Q.Insert ; InitNew(Q) ;
  Q.FindField('S_SECTION').AsString:=Sect ;
  Q.FindField('S_AXE').AsString:=LeCrit.E.Axe ;
  Q.FindField('S_LIBELLE').AsString:='Pour Solde (Situation)' ;
  Q.FindField('S_ABREGE').AsString:='Situation' ;
  Q.FindField('S_CONFIDENTIEL').AsString:='1' ;
  For i:=0 To 9 Do Q.FindField('S_TABLE'+IntToStr(i)).AsString:=TL[i] ;
  Q.Post ;
{$ENDIF}
END ;

Ferme(Q) ;
END ;

Function QuelTypeEcr(Var LeCrit : tCritGenereCERIC) : String ;
Var St : String ;
BEGIN
// FQ 18995 : ne prendre en compte que les écritures réellement choisies
//St:=' Y_QUALIFPIECE="N" ';
If LeCrit.E.OkNormale Then St:=St+' OR Y_QUALIFPIECE="N" ' ;
If LeCrit.E.OkSim Then St:=St+' OR Y_QUALIFPIECE="S" ' ;
If LeCrit.E.OkSit Then St:=St+' OR Y_QUALIFPIECE="U" ' ;
If LeCrit.E.OkPrev Then St:=St+' OR Y_QUALIFPIECE="P" ' ;
If LeCrit.E.OkRev Then St:=St+' OR Y_QUALIFPIECE="R" ' ;
//YMO 28/04/2006 FQ17955 Rajout IAS/IFRS
If LeCrit.E.OkIAIF Then St:=St+' OR Y_QUALIFPIECE="I" ' ;
if St <> '' then Result:=' AND ('+Copy(st,5, Length(St)-5)+') '
else Result := '';
END ;

Function FaitSelect(Racine : String ; Var LeCrit : tCritGenereCERIC ; Sens : String) : String ;
Var
 St, Where, Rub, Cpte1, Cpte2,
 Exclu1, Exclu2 : String ;
 fb1          : TFichierBase ;
 Q            : TQuery ;
 i            : integer ;

BEGIN
// Recherche de la rubrique
// YMO 03/2006 Remplacement de la table libre par une rubrique
If sens = 'C' then
  Rub:='"'+LeCrit.G.GTABLEC+'"'
else
  Rub:='"'+LeCrit.G.GTABLEP+'"';
Q:=OpenSQL('SELECT RB_COMPTE1, RB_COMPTE2, RB_EXCLUSION1, RB_EXCLUSION2 FROM RUBRIQUE'
          +' WHERE RB_RUBRIQUE = '+Rub, TRUE) ;
While Not Q.Eof Do
  BEGIN
  Cpte1:=Q.Findfield('RB_COMPTE1').AsString;
  Exclu1:=Q.Findfield('RB_EXCLUSION1').AsString;
  Cpte2:=Q.Findfield('RB_COMPTE2').AsString;
  Exclu2:=Q.Findfield('RB_EXCLUSION2').AsString;
  Q.Next ;
  END ;
Ferme(Q) ;

St:='SELECT Y_SECTION AS SECT, SUM(Y_DEBIT) AS DP, SUM(Y_CREDIT) AS CP FROM ANALYTIQ ' ;
{FQ18244 YMO 15/06/2006 La jointure ne tenait pas compte de l'axe,
doù fusion malvenue des montants de deux sections qui ont le même nom sur deux axes différents}
St:=St+' LEFT JOIN SECTION ON Y_SECTION=S_SECTION AND Y_AXE=S_AXE ' ;
St:=St+' LEFT JOIN GENERAUX ON Y_GENERAL=G_GENERAL ' ;
St:=St+' WHERE Y_SECTION LIKE "'+Racine+'%" AND Y_DATECOMPTABLE>="'+UsDateTime(LeCrit.E.DateE1)    
+'" AND Y_DATECOMPTABLE<="'+UsDateTime(LeCrit.E.DateE2)
+'" AND S_AFFAIREENCOURS="X" AND S_FINCHANTIER>"'+UsDateTime(LeCrit.E.DateE2)+'" ' ;
If LeCrit.E.Exo<>'' Then St:=St+' AND Y_EXERCICE="'+LeCrit.E.Exo+'" ' ;

// restrictuion sur tabkles libres généraux
For i:=0 To 9 Do
  If LeCrit.E.TLG[i]<>'' Then St:=St+' AND G_TABLE'+IntToStr(i)+'="'+LeCrit.E.TLG[i]+'" ' ;

// Restrictions définies au niveau des rubriques
fb1:=fbgene;
Where:=AnalyseCompte(Cpte1,fb1,False,False,False) ;
   if Where<>'' then St:=St+' AND '+Where ;
Where:=AnalyseCompte(Exclu1,fb1,True,False,False) ;
   if Where<>'' then St:=St+' AND '+Where ;
Where:=AnalyseCompte(Cpte2,fb1,False,False,False) ;
   if Where<>'' then St:=St+' AND '+Where ;
Where:=AnalyseCompte(Exclu2,fb1,True,False,False) ;
   if Where<>'' then St:=St+' AND '+Where ;

//St:=St+' AND ('+LeCrit.G.NomChamptable+'="'+LeCrit.G.GTABLEC+'" OR '+LeCrit.G.NOMChampTABLE+'="'+LeCrit.G.GTABLEP+'") ' ;
St:=St+' AND S_DEBCHANTIER<>"'+UsDateTime(IDate1900)+'" AND S_FINCHANTIER>"'+UsDateTime(LeCrit.G.DateG)+'" ' ;
//If not((LeCrit.E.OkSim) and (LeCrit.E.OkSit) and (LeCrit.E.OkPrev)
// and (LeCrit.E.OkRev) and (LeCrit.E.OkIAIF)) then
St:=St+QuelTypeEcr(LeCrit) ;
St:=St+' GROUP BY Y_SECTION';
//, '+LeCrit.G.NOMChampTABLE+' ' ;
Result:=St ;
END ;

Function FaitSelectSQLFIC(Var LeCrit : tCritGenereCERIC) : String ;
Var St : String ;
BEGIN
St:='SELECT Y_SECTION AS SECT, Y_GENERAL AS TL, SUM(Y_DEBIT) AS DP, SUM(Y_CREDIT) AS CP FROM ANALYTIQ ' ;
St:=St+' LEFT JOIN SECTION ON Y_SECTION=S_SECTION LEFT JOIN GENERAUX ON Y_GENERAL=G_GENERAL ' ;
St:=St+' WHERE Y_DATECOMPTABLE>="'+UsDateTime(LeCrit.E.DateE1)+'" AND Y_DATECOMPTABLE<="'+UsDateTime(LeCrit.E.DateE2)+'" ' ;
If LeCrit.E.Exo<>'' Then St:=St+' AND Y_EXERCICE="'+LeCrit.E.Exo+'" ' ;
//St:=St+' AND ('+LeCrit.G.NOMChampTABLE+'="'+LeCrit.G.GTABLEP+'") ' ;
St:=St+' AND S_DEBCHANTIER<>"'+UsDateTime(IDate1900)+'" AND S_FINCHANTIER>"'+UsDateTime(LeCrit.G.DateG)+'" ' ;
//If not((LeCrit.E.OkSim) and (LeCrit.E.OkSit) and (LeCrit.E.OkPrev)
//and (LeCrit.E.OkRev) and (LeCrit.E.OkIAIF)) then
St:=St+QuelTypeEcr(LeCrit) ;
St:=St+' GROUP BY Y_SECTION, Y_GENERAL' ;
Result:=St ;
END ;

Function QuelNum(Var LeCrit : tCritGenereCERIC) : Integer ;
Var Facturier : String ;
    NumP : Integer ;
BEGIN
Facturier:=LeCrit.G.SoucheN  ;
If LeCrit.G.QualG<>'N' Then Facturier:=LeCrit.G.SoucheS ;
SetIncNum(EcrGen,Facturier,NumP,LeCrit.G.DateG) ;
Result:=NumP ;
END ;

Procedure AlimTot(var Tot : TabDC4 ; D,C,DE,CE : Double) ;
BEGIN
Tot[1].TotDebit:=Arrondi(Tot[1].TotDebit+D,V_PGI.OkDecV) ;
Tot[1].TotCredit:=Arrondi(Tot[1].TotCredit+C,V_PGI.OkDecV) ;
Tot[2].TotDebit:=Arrondi(Tot[2].TotDebit+DE,V_PGI.OkDecE) ;
Tot[2].TotCredit:=Arrondi(Tot[2].TotCredit+CE,V_PGI.OkDecE) ;
END ;

Procedure RetoucheM(Var D,C,DE,CE : Double ; Charge : Boolean) ;
Var D1,C1,DE1,CE1 : Double ;
BEGIN
C1:=C ; D1:=D ; CE1:=CE ; DE1:=DE ;
If Charge Then
  BEGIN
//  If D=0 Then BEGIN D:=-C ; DE:=-CE ; C:=0 ; CE:=0 ; END ;
  D:=D1-C1 ; C:=0 ; DE:=DE1-CE1 ; CE:=0 ;
  END else
  BEGIN
//  If C=0 Then BEGIN C:=-D ; CE:=-DE ; D:=0 ; DE:=0 ; END ;
  C:=C1-D1 ; D:=0 ; CE:=CE1-DE1 ; DE:=0 ;
  END ;
END ;

Procedure MAJECR(QE : TQuery ; Var LeCrit : tCritGenereCERIC ; NumP,NumL : Integer ; D,C,DE,CE : Double ; OkC : Boolean) ;
Var LeCpt : String ;
BEGIN
If OkC Then BEGIN If NumL=1 Then LeCpt:=LeCrit.G.CptCB Else LeCpt:=LeCrit.G.CptCG ; END
       Else BEGIN If NumL=1 Then LeCpt:=LeCrit.G.CptPB Else LeCpt:=LeCrit.G.CptPG ; END ;
{$IFDEF EAGLCLIENT}
QE := TOB.Create('ECRITURE',nil,-1) ;

QE.PutValue('E_GENERAL',LeCpt) ;
QE.PutValue('E_DATECOMPTABLE',LeCrit.G.DateG) ;
QE.PutValue('E_NUMEROPIECE',NumP) ;
QE.PutValue('E_NUMLIGNE' ,NumL) ;

If NumL=2 Then QE.PutValue('E_ANA','X') ;
QE.PutValue('E_NUMECHE',0) ;
QE.PutValue('E_EXERCICE',LeCrit.G.ExoG) ;
QE.PutValue('E_DEBIT',D) ;
QE.PutValue('E_CREDIT',C) ;
QE.PutValue('E_DEBITDEV',D) ;
QE.PutValue('E_CREDITDEV',C) ;

//DE := D;
//CE := C;

QE.PutValue('E_REFINTERNE','Situation au '+DateToStr(LeCrit.G.DateG)) ;
//If NumL=2 Then //YMO 21/06/2006 Libellé sur les 2 lignes
QE.PutValue('E_LIBELLE',LeCrit.G.Lib) ;
QE.PutValue('E_NATUREPIECE','OD') ;
QE.PutValue('E_QUALIFPIECE',LeCrit.G.QualG) ;
QE.PutValue('E_ETABLISSEMENT',VH^.EtablisDefaut) ;
QE.PutValue('E_DEVISE',V_PGI.DevisePivot) ;
QE.PutValue('E_TAUXDEV',1) ;
QE.PutValue('E_JOURNAL',LeCrit.G.JalG) ;
QE.PutValue('E_ECRANOUVEAU','N') ;
QE.PutValue('E_CREERPAR','CER') ;
QE.PutValue('E_QUALIFORIGINE','GEN') ;
QE.PutValue('E_ETATLETTRAGE','RI') ;
QE.InsertDB(nil,False);

If LeCrit.G.QualG='N' Then MajSoldesEcritureTOB(QE,false) ;
{$ELSE}

QE.Insert ; InitNew(QE) ;
(* CA - FQ 18993 - Ce code doit être exécutée en Web Access et 2 tiers. Sinon, LeCpt n'est pas bon !!!
==> déplacé au début de la procédure
If OkC Then BEGIN If NumL=1 Then LeCpt:=LeCrit.G.CptCB Else LeCpt:=LeCrit.G.CptCG ; END
       Else BEGIN If NumL=1 Then LeCpt:=LeCrit.G.CptPB Else LeCpt:=LeCrit.G.CptPG ; END ;
       *)
QE.FindField('E_GENERAL').AsString:=LeCpt ;
QE.FindField('E_DATECOMPTABLE').AsDateTime:=LeCrit.G.DateG ;
QE.FindField('E_NUMEROPIECE').AsInteger:=NumP ;
QE.FindField('E_NUMLIGNE').AsInteger:=NumL ;

If NumL=2 Then QE.FindField('E_ANA').AsString:='X' ;
QE.FindField('E_NUMECHE').AsInteger:=0 ;
QE.FindField('E_EXERCICE').AsString:=LeCrit.G.ExoG ;
QE.FindField('E_DEBIT').AsFloat:=D ;
QE.FindField('E_CREDIT').AsFloat:=C ;
QE.FindField('E_DEBITDEV').AsFloat:=D ;
QE.FindField('E_CREDITDEV').AsFloat:=C ;

//DE := D;
//CE := C;

QE.FindField('E_REFINTERNE').AsString:='Situation au '+DateToStr(LeCrit.G.DateG) ;
//If NumL=2 Then //YMO 21/06/2006 Libellé sur les 2 lignes
QE.FindField('E_LIBELLE').AsString:=LeCrit.G.Lib ;
QE.FindField('E_NATUREPIECE').AsString:='OD' ;
QE.FindField('E_QUALIFPIECE').AsString:=LeCrit.G.QualG ;
QE.FindField('E_ETABLISSEMENT').AsString:=VH^.EtablisDefaut ;
QE.FindField('E_DEVISE').AsString:=V_PGI.DevisePivot ;
QE.FindField('E_TAUXDEV').AsFloat:=1 ;
QE.FindField('E_JOURNAL').AsString:=LeCrit.G.JalG ;
QE.FindField('E_ECRANOUVEAU').AsString:='N' ;
QE.FindField('E_CREERPAR').AsString:='CER' ;
QE.FindField('E_QUALIFORIGINE').AsString:='GEN' ;
QE.FindField('E_ETATLETTRAGE').AsString:='RI' ;
QE.Post ;

If LeCrit.G.QualG='N' Then MajSoldeCompte(QE) ;

{$ENDIF}


END ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/06/2006
Modifié le ... :   /  /
Description .. : Création de la ligne de contrepartie
Mots clefs ... :
*****************************************************************}
Procedure MAJLANA(ListeAna : TList ; Var LeCrit : tCritGenereCERIC ; NumP,NumV : Integer ; Sect : String ; D,C : Double ; OkC : Boolean ; AxePrinc : String) ;
Var O : TOBM ;
    OkOk : Boolean ;
    Axe, CptGene, SectionAtt : String ;
    Q1, Q2     : Tquery;
    i  : integer;
//    OrderNum : integer;
BEGIN
OkOk:=TRUE ;
If ListeAna.Count>=1 Then OkOk:=FALSE ;
If Not OkOk Then Exit ;
If OkC Then CptGene:=LeCrit.G.CptCG Else CptGene:=LeCrit.G.CptPG ;

Q1 := OpenSql ('SELECT * FROM GENERAUX WHERE G_GENERAL="'+CptGene+'"', True);
// FQ18243 YMO 15/06/2006 Alimentation de tous les axes des comptes de génération
for i:=1 to MAXAXE do
begin
Axe := 'A'+IntToStr(i);
if Q1.FindField ('G_VENTILABLE'+IntToStr(i)).asstring = 'X' then
begin
    If Axe=AxePrinc then
         SectionAtt:=Sect
    else
    begin
         Q2 := OpenSql ('SELECT * FROM Axe WHERE X_AXE="'+Axe+'"', True);
         if not Q2.EOF then
         begin
              if Q2.FindField ('X_SECTIONATTENTE').asstring <> '' then
                   Sectionatt := Q2.FindField ('X_SECTIONATTENTE').asstring
              else
                   Sectionatt :='ATTENTE'+Q2.findField ('X_AXE').asstring;

         end;
         ferme (Q2);
    end;

    {YMO  FQ18243 19/06/2006 Les numéros de ligne étaient incrémentés dans l'ordre des axes.
                             Il faut mettre l'axe traité en premier.
    OrderNum:=i;

    If Axe=AxePrinc then
      OrderNum:=1
    else
      if Axe<AxePrinc then
          OrderNum:=i+1;
    // On commence à 2
    Inc(OrderNum);     }


    O:=TOBM.Create(EcrAna,'',False) ;
    O.PutMvt('Y_GENERAL',CptGene);
    O.PutMvt('Y_AXE',Axe) ;
    O.PutMvt('Y_DATECOMPTABLE',LeCrit.G.DateG) ;
    O.PutMvt('Y_NUMEROPIECE',NumP) ;
    O.PutMvt('Y_NUMLIGNE', 2) ;     // YMO 21/06/2006 numéro de ligne 2 sur tous les axes
    O.PutMvt('Y_SECTION',SectionAtt) ;
    O.PutMvt('Y_EXERCICE',LeCrit.G.ExoG) ;
    O.PutMvt('Y_DEBIT',D) ;
    O.PutMvt('Y_CREDIT',C) ;
    O.PutMvt('Y_DEBITDEV',D) ;
    O.PutMvt('Y_CREDITDEV',C) ;
    O.PutMvt('Y_REFINTERNE','Situation au '+DateToStr(LeCrit.G.DateG)) ;
    O.PutMvt('Y_AFFAIRE','AUTOCUTOFF') ;
    O.PutMvt('Y_LIBELLE',LeCrit.G.Lib) ;
    O.PutMvt('Y_NATUREPIECE','OD') ;
    O.PutMvt('Y_QUALIFPIECE',LeCrit.G.QualG) ;
    O.PutMvt('Y_TYPEANALYTIQUE','-') ;
    O.PutMvt('Y_ETABLISSEMENT',VH^.EtablisDefaut) ;
    O.PutMvt('Y_DEVISE',V_PGI.DevisePivot) ;
    O.PutMvt('Y_TAUXDEV',1) ;
    O.PutMvt('Y_JOURNAL',LeCrit.G.JalG) ;
    O.PutMvt('Y_NUMVENTIL',NumV) ;
    O.PutMvt('Y_ECRANOUVEAU','N') ;
    O.PutMvt('Y_CREERPAR','ACO') ;

    ListeAna.Add(O) ;
end;
end;
ferme (Q1);
END ;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/06/2006
Modifié le ... :   /  /
Description .. : Ise à jour des montants
Mots clefs ... :
*****************************************************************}
Procedure RetoucheLAna(L : TList ; Tot : TabDC4 ; Var LeCrit : tCritGenereCERIC) ;
Var i : Integer ;
    O : TOBM ;
    T,M : Double ;
BEGIN
For i:=0 To L.Count-1 Do
  BEGIN
  O:=TOBM(L[i]) ; If O=NIL Then Continue ;
  T:=Arrondi(Tot[1].TotDebit+Tot[1].TotCredit,V_PGI.OkDecV) ;

  O.PutMvt('Y_DEBIT',Tot[1].TotDebit) ;
  O.PutMvt('Y_CREDIT',Tot[1].TotCredit) ;
  O.PutMvt('Y_DEBITDEV',Tot[1].TotDebit) ;
  O.PutMvt('Y_CREDITDEV',Tot[1].TotCredit) ;

  O.PutMvt('Y_TOTALECRITURE',T) ;
  O.PutMvt('Y_TOTALDEVISE',T) ;

  M:=O.GetMvt('Y_DEBIT')+O.GetMvt('Y_CREDIT') ;
  If T<>0 Then O.PutMvt('Y_POURCENTAGE',Arrondi((100*M)/T,5)) ;
  END ;
END ;

Procedure MajAna(L : TList ; QA : TQuery ; Var LeCrit : tCritGenereCERIC) ;
Var i : Integer ;
    O : TOBM ;
BEGIN
InitMove(L.Count-1,'') ;
For i:=0 To L.Count-1 Do
  BEGIN
  MoveCur(False) ;

  // FQ18282 YMO 15/06/2006 Uniformisation du code 2T/CWAS
  O:=TOBM(L[i]) ; if O=Nil then Continue ;
  O.InsertDB(nil,False);
  If LeCrit.G.QualG='N' Then
      MajSoldeSectionTOB(O,True);

  END ;
FiniMove ;
END ;

Function Genere1Mvt(Racine : String ; Var LeCrit : tCritGenereCERIC ; QE,QA : TQuery ; Var CritExpMvt : TCritExpMvt ; Sens : String) : Integer ;
Var St : String ;
    NumVC,NumVP : Integer ;
    Sect : String ;
    ListeAnaC,ListeAnaP : TList ;
    Q : TQuery ;
    D,C,DE,CE : Double ;
    TotCha,TotPro : TabDC4 ;
    NumPC,NumPP,NbGenere : Integer ;
    PremFoisC,PremFoisP : Boolean ;
    TL : TTL ;
BEGIN
NbGenere:=0 ;

ListeAnaC:=TList.Create ;
ListeAnaP:=TList.Create ;
// Sélectionne toutes les écritures analytique correspondant aux sections ouvertes,
// uniquement les sections correspondant aux racines spécifiées sur les rubriques
St:=FaitSelect(Racine,LeCrit,Sens) ;

LeCrit.G.SectProv:=Racine;
Q:=OpenSQL(St,TRUE) ;
InitMove(RecordsCount(Q),'') ;
Fillchar(TotCha,SizeOf(TotCha),#0) ; Fillchar(TotPro,SizeOf(TotPro),#0) ;

PremFoisC:=TRUE ; PremFoisP:=TRUE ; NumVP:=0 ; NumVC:=0 ; NumPC:=0; NumPP:=0;
While Not Q.Eof Do
  BEGIN
  D:=Arrondi(Q.Findfield('DP').AsFloat,V_PGI.OkDecV) ; C:=Arrondi(Q.Findfield('CP').AsFloat,V_PGI.OkDecV) ;

  Sect:=Q.Findfield('SECT').AsString ;
  If PremFoisC And PremFoisP Then LeCrit.G.Lib:=QuelLib(Sect,TL) ;
  Sect:=LeCrit.G.SectProv ;

  If Sens='C' then
  BEGIN
    RetoucheM(D,C,DE,CE,TRUE) ;
    If PremFoisC And ((D<>0) Or (C<>0)) Then BEGIN NumPC:=QuelNum(LeCrit) ; PremFoisC:=FALSE ; END ;
    AlimTot(TotCha,C,D,CE,DE) ;
    If (D<>0) Or (C<>0) Then
      BEGIN
      Inc(NumVC) ;
      // Création de l'écriture analytique charges
      MajLANA(ListeAnaC,LeCrit,NumPC,NumVC,Sect,C,D,TRUE, Lecrit.E.Axe) ;

      END ;
  END Else
  If Sens='P' then
  BEGIN
    RetoucheM(D,C,DE,CE,FALSE) ;
    If PremFoisP And ((D<>0) Or (C<>0)) Then BEGIN NumPP:=QuelNum(LeCrit) ; PremFoisP:=FALSE ; END ;
    AlimTot(TotPro,C,D,CE,DE) ;
    If (D<>0) Or (C<>0) Then
      BEGIN
      Inc(NumVP) ;
      // Création de l'écriture analytique produits
      MajLANA(ListeAnaP,LeCrit,NumPP,NumVP,Sect,C,D,FALSE, Lecrit.E.Axe) ;

      END ;
  END ;
  MoveCur(FALSE) ;
  Q.Next ;
  END ;
Ferme(Q) ;
If ListeANAC.Count>0 Then
  BEGIN
  RetoucheLANA(ListeANAC,TotCha,LeCrit) ;
  MajAna(ListeANAC,QA,LeCrit) ;
  // Création des écritures bilan-gestion
  MAJECR(QE,LeCrit,NumPC,2,TotCHA[1].TotDebit,TotCHA[1].TotCredit,TotCHA[2].TotDebit,TotCHA[2].TotCredit,TRUE) ;
  MAJECR(QE,LeCrit,NumPC,1,TotCHA[1].TotCredit,TotCHA[1].TotDebit,TotCHA[2].TotCredit,TotCHA[2].TotDebit,TRUE) ;
  CritExpMvt.OkExport:=TRUE ;

  If NumPC < CritExpMvt.NumP1 Then
      CritExpMvt.NumP1:=NumPC ;

  If NumPC > CritExpMvt.NumP2 Then
      CritExpMvt.NumP2:=NumPC ;

  Inc(NbGenere) ;
  END ;
If ListeANAP.Count>0 Then
  BEGIN
  RetoucheLANA(ListeANAP,TotPro,LeCrit) ;
  MajAna(ListeANAP,QA,LeCrit) ;
  MAJECR(QE,LeCrit,NumPP,2,TotPro[1].TotDebit,TotPro[1].TotCredit,TotPro[2].TotDebit,TotPro[2].TotCredit,FALSE) ;
  MAJECR(QE,LeCrit,NumPP,1,TotPro[1].TotCredit,TotPro[1].TotDebit,TotPro[2].TotCredit,TotPro[2].TotDebit,FALSE) ;
  CritExpMvt.OkExport:=TRUE ;
  If NumPC<CritExpMvt.NumP1 Then CritExpMvt.NumP1:=NumPP ;
  If NumPC>CritExpMvt.NumP2 Then CritExpMvt.NumP2:=NumPP ;
  Inc(NbGenere) ;
  END ;
If (ListeANAC.Count>0) Or (ListeANAP.Count>0) Then VerifSectProv(LeCrit,TL) ;
FiniMove ;
VideListe(ListeAnaC) ; ListeAnaC.Free ;
VideListe(ListeAnaP) ; ListeAnaP.Free ;
Result:=NbGenere ;
END ;

Function TFGenere.GenereMvt(L : tStringList ; Var LeCrit : tCritGenereCERIC ; Var CritExpMvt : TCritExpMvt ; Sens : String) : Integer ;
Var i,j,NbGenere : Integer ;
    QE,QA : TQuery ;
BEGIN
Result:=0 ; NbGenere:=0 ;
If L.Count<=0 Then Exit ;
QE:=OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+W_W+'"',FALSE) ; QE.Open ;
QA:=OpenSQL('SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="'+W_W+'"',FALSE) ; QA.Open ;
For i:=0 To L.Count-1 Do
  BEGIN
  LEnCours.Caption:=HM.Mess[13]+' '+L[i] ; Application.ProcessMessages ;
  j:=Genere1Mvt(L[i],LeCrit,QE,QA,CritExpMvt,Sens) ;
  NbGenere:=NbGenere+j ;
  END ;
Ferme(QE) ; Ferme(QA) ;
Result:=NbGenere ;
END ;

Procedure InitCritExpMvt(Var LeCrit : tCritGenereCERIC ; Var CritExpMvt : TCritExpMvt) ;
BEGIN
With LeCrit,CritExpMvt Do
  BEGIN
  Format:='CGN' ; Exo:=G.ExoG ; Etab:=VH^.EtablisDefaut ; Jal1:=G.JalG ; Jal2:=G.JalG ;
  TypeEcr:=G.QualG ; Date1:=G.DateG ; Date2:=G.DateG ; NomFic:=G.NomFic ;
  NumP1:=999999999 ; NumP2:=0 ;
  END ;
END ;

Procedure FAitFicSQL(Var LeCrit : tCritGenereCERIC) ;
Var LeNom ,St : String ;
    F : TextFile ;
BEGIN
St:=FaitSelectSQLFIC(LeCrit) ;
LeNom:=NewNomFic(LeCrit.G.NomFic,'SQL') ;
Assign(F,LeNom) ;
Rewrite(F) ; WriteLn(F,St) ; Close(F) ;
END ;

procedure TFGenere.BValiderClick(Sender: TObject);
Var LRacine : TStringList ;
    CritExpMvt : TCritExpMvt ;
    NbGenere, NbGenereC, NbGenereP : integer ;
begin
If HM.Execute(11,Caption,'')<>mrYes Then Exit ;
RecupCrit ;
If CritOk(CritG) Then
  BEGIN
  Fillchar(CritExpMvt,SizeOf(CritExpMvt),#0) ;
  InitCritExpMvt(CritG,CritExpMvt) ;
  EnableControls(Self,FALSE) ; LEnCours.Visible:=TRUE ; LEnCours.Caption:='' ;
  LRacine:=TStringList.Create ; LRacine.Sorted:=TRUE ; LRacine.Duplicates:=DupIgnore ;
  // Recense toutes les sections remplissant les conditions sélectionnées à l'écran
  // qui n'ont pas de date de fin de chantier (ouvertes)
  FaitListeRacine(LRacine,CritG) ;

  NbGenereC:=GenereMvt(LRacine,CritG,CritExpMvt,'C') ;
  NbGenereP:=GenereMvt(LRacine,CritG,CritExpMvt,'P') ;
  LRacine.Clear ; LRacine.Free ;
  EnableControls(Self,TRUE) ; LEnCours.Visible:=FALSE ;
  HM.Execute(12,Caption,'') ;
  NbGenere:=NbGenereC+NbGenereP;
 { YMO 04/2006 Pas d'export fichier
  If (NbGenere>0) And CritExpMvt.OkExport And (CritExpMvt.NomFic<>'') Then
    BEGIN
    LEnCours.Caption:=HM.Mess[20] ; LEnCours.Visible:=TRUE ;
  //  ExternalExportMvt('FEC',CritExpMvt) ; FaitFicSQL(CritG) ;
    LEnCours.Visible:=FALSE ; LEnCours.Caption:='' ;
    END ;
}
  If NbGenere<=0 Then HM.Execute(19,'','') ;
  END ;
end;

(*
Procedure MAJANA(QA : TQuery ; Var LeCrit : tCritGenereCERIC ; NumP,NumV : Integer ; Sect : String ; D,C,DE,CE : Double ; OkC : Boolean) ;
BEGIN
QA.Insert ; InitNew(QA) ;
If OkC Then QA.FindField('Y_GENERAL').AsString:=LeCrit.G.CptCG
       Else QA.FindField('Y_GENERAL').AsString:=LeCrit.G.CptPG ;
QA.FindField('Y_AXE').AsString:=LeCrit.E.Axe ;
QA.FindField('Y_DATECOMPTABLE').AsDateTime:=LeCrit.G.DateG ;
QA.FindField('Y_NUMEROPIECE').AsInteger:=NumP ;
QA.FindField('Y_NUMLIGNE').AsInteger:=2 ;
QA.FindField('Y_SECTION').AsString:=Sect ;
QA.FindField('Y_EXERCICE').AsString:=LeCrit.G.ExoG ;
QA.FindField('Y_DEBIT').AsFloat:=C ;
QA.FindField('Y_CREDIT').AsFloat:=D ;
QA.FindField('Y_DEBITDEV').AsFloat:=C ;
QA.FindField('Y_CREDITDEV').AsFloat:=D ;
QA.FindField('Y_DEBITEURO').AsFloat:=CE ;
QA.FindField('Y_CREDITEURO').AsFloat:=DE ;
QA.FindField('Y_REFINTERNE').AsString:='Situation au '+DateToStr(LeCrit.G.DateG) ;
QA.FindField('Y_LIBELLE').AsString:=LeCrit.G.Lib ;
QA.FindField('Y_NATUREPIECE').AsString:='OD' ;
QA.FindField('Y_QUALIFPIECE').AsString:=LeCrit.G.QualG ;
QA.FindField('Y_TYPEANALYTIQUE').AsString:='-' ;
QA.FindField('Y_ETABLISSEMENT').AsString:=VH^.EtablisDefaut ; ;
QA.FindField('Y_DEVISE').AsString:=V_PGI.DevisePivot ;
QA.FindField('Y_TAUXDEV').AsFloat:=1 ;
QA.FindField('Y_JOURNAL').AsString:=LeCrit.G.JalG ;
QA.FindField('Y_NUMVENTIL').AsString:=NumV ;
QA.FindField('Y_ECRANOUVEAU').AsString:='N' ;
QA.FindField('Y_CREERPAR').AsString:='CER' ;
QA.FindField('Y_SAISIEEURO').AsString:='-' ;
QA.Post ;
END ;
*)

procedure TFGenere.HMTradBeforeTraduc(Sender: TObject);
begin
LibellesTableLibre(PzLibre,'TS_TABLE','S_TABLE','S') ;
LibellesTableLibre(GenTLibre,'TG_TABLE','G_TABLE','G') ;
end;

procedure TFGenere.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  //YMO 15/05/2006 FQ18127 Nouvelle gestion des filtres
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'MVTSECT');

end;

procedure TFGenere.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  ObjFiltre.Free;
end;

procedure TFGenere.S_SECTIONExit(Sender: TObject);
begin
  if (Trim(THEdit(Sender).Text) = '') or (TestJoker(THEdit(Sender).Text)) then Exit;
  if Length(THEdit(Sender).Text) < VH^.Cpta[AxeToFb(S_AXE.Value)].Lg then
    THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, AxeToFb(S_AXE.Value));
  if (THEdit(Sender).Name = 'S_SECTION') and (S_SECTION_.Text='') then S_SECTION_.Text :=S_SECTION.Text;  
end;

end.
