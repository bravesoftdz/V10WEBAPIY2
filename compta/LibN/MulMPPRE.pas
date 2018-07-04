unit MulMPPRE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB, DBTables, Hqry,
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, Saisie, SaisUtil,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil,
  EncUtil, LettUtil, GenereMP, UTOB
{$IFNDEF IMP}
  ,SaisBor, HRichOLE
{$ENDIF}
  ;

procedure GenerePrelevements ;

type
  TFMulMPPRE = class(TFMul)
    HM: THMsgBox;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE9: THCpteEdit;
    TE_NUMEROPIECE: THLabel;
    HLabel2: THLabel;
    TE_DEBIT: TLabel;
    TE_DEBIT_: TLabel;
    TE_CREDIT: TLabel;
    TE_CREDIT_: TLabel;
    TE_ETABLISSEMENT: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    E_NUMEROPIECE_: THCritMaskEdit;
    E_DEBIT: THCritMaskEdit;
    E_DEBIT_: THCritMaskEdit;
    E_CREDIT: THCritMaskEdit;
    E_CREDIT_: THCritMaskEdit;
    E_ETABLISSEMENT: THValComboBox;
    PEcritures: TTabSheet;
    Bevel6: TBevel;
    TE_JOURNAL: THLabel;
    TE_NATUREPIECE: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    TE_EXERCICE: THLabel;
    TE_DATEECHEANCE: THLabel;
    TE_DATEECHEANCE2: THLabel;
    E_EXERCICE: THValComboBox;
    E_JOURNAL: THValComboBox;
    E_NATUREPIECE: THValComboBox;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_NUMECHE: THCritMaskEdit;
    XX_WHEREAN: TEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    E_ECHE: THCritMaskEdit;
    E_ETATLETTRAGE: THCritMaskEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHEREENC: TEdit;
    XX_WHEREVIDE: TEdit;
    E_DATEECHEANCE: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    cFactCredit: TCheckBox;
    HLabel4: THLabel;
    E_GENERAL: THCritMaskEdit;
    HLabel17: THLabel;
    E_DEVISE: THValComboBox;
    HLabel7: THLabel;
    CATEGORIE: THValComboBox;
    Label14: TLabel;
    E_MODEPAIE: THValComboBox;
    H_ModeSaisie: TLabel;
    CModeSaisie: THValComboBox;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCritMaskEdit;
    FSelectAll: TCheckBox;
    HDiv: THMsgBox;
    XX_WHEREMP: TEdit;
    BParams: TToolbarButton97;
    HLabel1: THLabel;
    E_NUMENCADECA: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CATEGORIEChange(Sender: TObject);
    procedure BParamsClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private    { Déclarations privées }
    ANouveau  : boolean ;
    procedure InitCriteres ;
    procedure InitPrelevements ;
    procedure PrechargeOrigines ;
    procedure MarqueOrigine ;
    procedure ToutMarquer ;
    procedure VireInutiles ;
    procedure ConstitueOrigines ;
    procedure InitCommuns ( TOBD : TOB ) ;
    procedure TripoteDetail ( TOBD : TOB ) ;
    Function  GetLeMontant ( Montant,Couv,Sais : Double ; tsm : TSorteMontant ) : Double ;
    procedure AffecteG ( TOBL : TOB ) ;
    procedure FinirDestination ( TOBR : TOB ) ;
    procedure ConstitueDest ;
  public
    ParmsMP : tGenereMP ;
    TOBORig,TOBDest : TOB ;   
  end;

implementation

{$R *.DFM}

procedure GenerePrelevements ;
var X : TFMulMPPRE ;
    PP : THPanel ;
begin
PP:=FindInsidePanel ;
X:=TFMulMPPRE.Create(Application) ;
X.FNomFiltre:='CPENCPRELEVEMENT' ;
X.Q.Manuel:=True ; X.Q.Liste:='CPENCPRELEVEMENT' ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

(*========================= METHODES DE LA FORM ==============================*)
procedure TFMulMPPRE.FormCreate(Sender: TObject);
begin
  inherited;
MemoStyle:=msBook ;
RempliSelectEuro(CModeSaisie,HDiv.Mess[0]) ;
TOBOrig:=TOB.Create('',Nil,-1) ;
TOBDest:=TOB.Create('',Nil,-1) ;
end;

procedure TFMulMPPRE.FormShow(Sender: TObject);
begin
InitCriteres ;
InitPrelevements ;
  inherited;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
XX_WHEREVIDE.Text:='' ;
end;

(*============================ INITIALISATIONS ===============================*)
procedure TFMulMPPRE.InitPrelevements ;
BEGIN
XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' ;
CATEGORIE.Value:='PRE' ;
CatToMP(CATEGORIE.Value,E_MODEPAIE.Items,E_MODEPAIE.Values,tslAucun,True) ;
E_GENERAL.DataType:='TZGENCAIS' ;
E_AUXILIAIRE.DataType:='TZTTOUTDEBIT' ;
FillChar(ParmsMP,Sizeof(ParmsMP),#0) ;
ParmsMP.NomFSelect:=FNomFiltre ;
ParmsMP.Cat:=CATEGORIE.Value ;
Caption:=HDiv.Mess[1] ; UpdateCaption(Self) ;
END ;

procedure TFMulMPPRE.InitCriteres ;
BEGIN
LibellesTableLibre(PzLibre,'TE_TABLE','E_TABLE','E') ;
E_EXERCICE.Value:=VH^.Entree.Code ;
E_DATECOMPTABLE.Text:=DateToStr(VH^.Entree.Deb) ; E_DATECOMPTABLE_.Text:=DateToStr(VH^.Entree.Fin) ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
E_DEVISE.Value:=V_PGI.DevisePivot ;
END ;

(*================================ CRITERES ==================================*)
procedure TFMulMPPRE.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFMulMPPRE.CATEGORIEChange(Sender: TObject);
Var St,sMode,sCat,Junk : String ;
    i : integer ;
begin
  inherited;
XX_WHEREMP.Text:='' ; if CATEGORIE.Value='' then Exit ;
for i:=0 to VH^.MPACC.Count-1 do
    BEGIN
    St:=VH^.MPACC[i] ;
    sMode:=ReadtokenSt(St) ; Junk:=ReadtokenSt(St) ; sCat:=ReadtokenSt(St) ;
    if sCat=CATEGORIE.Value then XX_WHEREMP.Text:=XX_WHEREMP.Text+' OR E_MODEPAIE="'+sMode+'"' ;
    END ;
if XX_WHEREMP.Text='' then XX_WHEREMP.Text:='E_MODEPAIE="aaa"' else // ne rien trouver
   BEGIN
   St:=XX_WHEREMP.Text ;
   Delete(St,1,4) ;
   XX_WHEREMP.Text:=St ;
   END ;
end;

procedure TFMulMPPRE.BParamsClick(Sender: TObject);
begin
  inherited;
ParamsMP(False,ParmsMP) ;
end;

procedure TFMulMPPRE.FListeDblClick(Sender: TObject);
Var sMode : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
sMode:=Q.FindField('E_MODESAISIE').AsString ;
if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q,TypeAction)
                                  else TrouveEtLanceSaisie(Q,taConsult,'N') ;
end;

procedure TFMulMPPRE.PrechargeOrigines ;
Var QQ : TQuery ;
    sWhere : String ;
BEGIN
TOBOrig.ClearDetail ;
QQ:=OpenSQL('SELECT * FROM ECRITURE '+RecupWhereCritere(Pages),True) ;
TOBOrig.LoadDetailDB('ECRITURE','','',QQ,False,True) ;
Ferme(QQ) ;
if TOBOrig.Detail.Count>0 then TOBOrig.Detail[0].AddChampSup('MARQUE',True) ;
END ;

procedure TFMulMPPRE.MarqueOrigine ;
Var TOBL : TOB ;
BEGIN
TOBL:=TOBOrig.FindFirst(['E_JOURNAL','E_EXERCICE','E_DATECOMPTABLE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                        [Q.FindField('E_JOURNAL').AsString,Q.FindField('E_EXERCICE').AsString,
                         Q.FindField('E_DATECOMPTABLE').AsDateTime,Q.FindField('E_NUMEROPIECE').AsInteger,
                         Q.FindField('E_NUMLIGNE').AsInteger,Q.FindField('E_NUMECHE').AsInteger],False) ;
if TOBL<>Nil then TOBL.PutValue('MARQUE','X') ;
END ;

procedure TFMulMPPRE.ToutMarquer ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBORig.Detail.Count-1 do
    BEGIN
    TOBL:=TOBOrig.Detail[i] ;
    TOBL.PutValue('MARQUE','X') ;
    END ;
END ;

procedure TFMulMPPRE.VireInutiles ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=TOBOrig.Detail.Count-1 downto 0 do
    BEGIN
    TOBL:=TOBOrig.Detail[i] ;
    if TOBL.GetValue('MARQUE')<>'X' then BEGIN TOBL.Free ; TOBL:=Nil ; END ;  
    END ;
END ;

procedure TFMulMPPRE.ConstitueOrigines ;
Var i : integer ;
BEGIN
PrechargeOrigines ;
if Not FListe.AllSelected then
   BEGIN
   for i:=0 to FListe.NbSelected-1 do
       BEGIN
       FListe.GotoLeBookmark(i) ;
       MarqueOrigine ;
       END ;
   END else
   BEGIN
   ToutMarquer ;
   END ;
VireInutiles ;
END ;

procedure TFMulMPPRE.InitCommuns ( TOBD : TOB ) ;
Var i : integer ;
BEGIN
// Lettrage
TOBD.PutValue('E_LETTRAGE','') ; TOBD.PutValue('E_ETATLETTRAGE','RI') ;
TOBD.PutValue('E_COUVERTURE',0) ; TOBD.PutValue('E_COUVERTUREDEV',0) ; TOBD.PutValue('E_COUVERTUREEURO',0) ;
TOBD.PutValue('E_LETTRAGEDEV','-') ; TOBD.PutValue('E_LETTRAGEEURO','-') ;
TOBD.PutValue('E_ETAT','0000000000') ;
// Eches + Anas
TOBD.PutValue('E_ECHE','-') ; TOBD.PutValue('E_ANA','-') ;
// Saisie
TOBD.PutValue('E_MODESAISIE','-') ;
TOBD.PutValue('E_SAISIEEURO','-') ;
TOBD.PutValue('E_QUALIFORIGINE','') ;
// Enc et Dec
TOBD.PutValue('E_SUIVDEC','') ;
TOBD.PutValue('E_NOMLOT','') ;
TOBD.PutValue('E_REFRELEVE','') ;
TOBD.PutValue('E_ENCAISSEMENT',SensEnc(TOBD.GetValue('E_DEBITDEV'),TOBD.GetValue('E_CREDITDEV'))) ;
// Tiers
TOBD.PutValue('E_DATERELANCE',iDate1900) ;
TOBD.PutValue('E_NIVEAURELANCE',0) ;
// TVA
for i:=1 to 4 do TOBD.PutValue('E_ECHEENC'+IntToStr(i),0) ;
TOBD.PutValue('E_ECHEDEBIT',0) ;
// Divers
TOBD.PutValue('E_VALIDE','-') ;
TOBD.PutValue('E_AVOIRRBT','-') ;
TOBD.PutValue('E_PIECETP','') ;
TOBD.PutValue('E_EQUILIBRE','-') ;
TOBD.PutValue('E_CFONBOK','-') ;
TOBD.PutValue('E_EDITEETATTVA','-') ;
END ;

procedure TFMulMPPRE.TripoteDetail ( TOBD : TOB ) ;
BEGIN
InitCommuns(TOBD) ;
if ParmsMP.Ref2<>'' then TOBD.PutValue('E_REFINTERNE',ParmsMP.Ref2) ;
if ParmsMP.Lib2<>'' then TOBD.PutValue('E_LIBELLE',ParmsMP.Lib2) ;
TOBD.PutValue('E_GENERAL',ParmsMP.CptG) ;
if ParmsMP.MPG<>'' then TOBD.PutValue('E_MODEPAIE',ParmsMP.MPG) ;
// Cas particuliers
if ParmsMP.GenColl then
   BEGIN
   TOBD.PutValue('E_ETATLETTRAGE','AL') ;
   TOBD.PutValue('E_ECHE','X') ;
   TOBD.PutValue('E_NUMECHE',1) ;
   END else
   BEGIN
   TOBD.PutValue('E_AUXILIAIRE','') ;
   if ParmsMP.GenLett then
      BEGIN
      TOBD.PutValue('E_ETATLETTRAGE','AL') ;
      TOBD.PutValue('E_ECHE','X') ;
      TOBD.PutValue('E_NUMECHE',1) ;
      END ;
   if ParmsMP.GenPoint then
      BEGIN
      TOBD.PutValue('E_ETATLETTRAGE','RI') ;
      TOBD.PutValue('E_ECHE','X') ;
      TOBD.PutValue('E_NUMECHE',1) ;
      END ;
   END ;
if ParmsMP.GenVent then
   BEGIN
   TOBD.PutValue('E_ANA','X') ;
   END ;
END ;

Function TFMulMPPRE.GetLeMontant ( Montant,Couv,Sais : Double ; tsm : TSorteMontant ) : Double ;
BEGIN
Result:=0 ;
if Arrondi(Montant,4)=0 then Exit ;
if Arrondi(Sais,4)=0 then
   BEGIN
   Result:=Arrondi(Montant-Couv,6) ;
   END else
   BEGIN
   if E_DEVISE.Value=V_PGI.DevisePivot then
      BEGIN
      if tsm=tsmPivot then Result:=Sais else Result:=0 ;
      END else
      BEGIN
      if tsm=tsmDevise then Result:=Sais else Result:=0 ;
      END ;
   END ;
END ;

procedure TFMulMPPRE.AffecteG ( TOBL : TOB ) ;
BEGIN
TOBL.PutValue('E_EXERCICE',QuelExoDT(ParmsMP.DateC)) ;
TOBL.PutValue('E_DATECOMPTABLE',ParmsMP.DateC) ;
TOBL.PutValue('E_JOURNAL',ParmsMP.JalG) ;
if ((TOBL.GetValue('E_ECHE')='X') and (ParmsMP.DateE>0)) then TOBL.PutValue('E_DATEECHEANCE',ParmsMP.DateE) ;
END ;

procedure TFMulMPPRE.FinirDestination ( TOBR : TOB ) ;
Var i,Ind,NewNum : integer ;
    TOBL  : TOB ;
    XD,XC : Double ;
BEGIN
for i:=0 to TOBDest.Detail.Count-1 do
    BEGIN
    TOBL:=TOBDest.Detail[i] ;
    XD:=GetLeMontant(TOBL.GetValue('E_DEBIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot) ;
    XC:=GetLeMontant(TOBL.GetValue('E_CREDIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot) ;
    TOBL.PutValue('E_DEBIT',XC) ; TOBL.PutValue('E_CREDIT',XD) ;

    XD:=GetLeMontant(TOBL.GetValue('E_DEBITDEV'),TOBL.GetValue('E_COUVERTUREDEV'),TOBL.GetValue('E_SAISIMP'),tsmDevise) ;
    XC:=GetLeMontant(TOBL.GetValue('E_CREDITDEV'),TOBL.GetValue('E_COUVERTUREDEV'),TOBL.GetValue('E_SAISIMP'),tsmDevise) ;
    TOBL.PutValue('E_DEBITDEV',XC) ; TOBL.PutValue('E_CREDITDEV',XD) ;

    XD:=GetLeMontant(TOBL.GetValue('E_DEBITEURO'),TOBL.GetValue('E_COUVERTUREEURO'),TOBL.GetValue('E_SAISIMP'),tsmEuro) ;
    XC:=GetLeMontant(TOBL.GetValue('E_CREDITEURO'),TOBL.GetValue('E_COUVERTUREEURO'),TOBL.GetValue('E_SAISIMP'),tsmEuro) ;
    TOBL.PutValue('E_DEBITEURO',XC) ; TOBL.PutValue('E_CREDITEURO',XD) ;
    if TOBL.GetValue('E_ETATLETTRAGE')<>'RI' then
       BEGIN
       TOBL.PutValue('E_LETTRAGE','') ;
       TOBL.PutValue('E_ETATLETTRAGE','AL') ;
       TOBL.PutValue('E_ECHE','X') ;
       TOBL.PutValue('E_NUMECHE',1) ;
       END ;
    if ParmsMP.Ref1<>'' then TOBL.PutValue('E_REFINTERNE',ParmsMP.Ref1) ;
    if ParmsMP.Lib1<>'' then TOBL.PutValue('E_LIBELLE',ParmsMP.Lib1) ;
    TOBL.PutValue('E_SAISIMP',0) ;
    TOBL.PutValue('E_COUVERTURE',0) ; TOBL.PutValue('E_COUVERTUREDEV',0) ; TOBL.PutValue('E_COUVERTUREEURO',0) ;
    END ;
for i:=TOBR.Detail.Count-1 downto 0 do
    BEGIN
    TOBL:=TOBR.Detail[i] ;
    Ind:=TOBL.GetValue('INDICE') ;
    TOBL.ChangeParent(TOBDest,Ind) ;
    END ;
NewNum:=GetNewNumJal(ParmsMP.JalG,True,ParmsMP.DateC) ;
for i:=0 to TOBDest.Detail.Count-1 do
    BEGIN
    TOBL:=TOBDest.Detail[i] ;
    TOBL.PutValue('E_NUMLIGNE',i+1) ;
    TOBL.PutValue('E_NUMEROPIECE',NewNum) ;
    AffecteG(TOBL) ;
    END ;
END ;

procedure TFMulMPPRE.ConstitueDest ;
Var TOBR,TOBL,TOBDETR : TOB ;
    i,Lasti : integer ;
    DP,CP,DE,CE,DD,CD,XD,XC : double ;
    TypG,OldAux,OldGen : String ;
    OldEche : TDateTime ;
    Rupt,Premier : boolean ;
BEGIN
TOBDest.Dupliquer(TOBOrig,True,True,False) ;
TOBDest.Detail.Sort('E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE') ;
TOBR:=TOB.Create('',Nil,-1) ;
DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ; DE:=0 ; CE:=0 ;
TypG:=ParmsMP.GroupeEncadeca ; Rupt:=False ; Premier:=True ;
for i:=0 to TOBDest.Detail.Count-1 do
    BEGIN
    TOBL:=TOBDest.Detail[i] ; Lasti:=i ;
    if Premier then
       BEGIN
       OldAux:=TOBL.GetValue('E_AUXILIAIRE') ;
       OldGen:=TOBL.GetValue('E_GENERAL') ;
       OldEche:=TOBL.GetValue('E_DATEECHEANCE') ;
       END ;
    if TypG='DET' then Rupt:=True else
     if TypG='GLO' then Rupt:=False else
      if TypG='AUX' then Rupt:=((OldAux<>TOBL.GetValue('E_AUXILIAIRE')) or (OldGen<>TOBL.GetValue('E_GENERAL'))) else
        if TypG='ECH' then Rupt:=((OldAux<>TOBL.GetValue('E_AUXILIAIRE')) or (OldGen<>TOBL.GetValue('E_GENERAL')) or (OldEche<>TOBL.GetValue('E_DATEECHEANCE'))) ;
    Premier:=False ;
    if Rupt then
       BEGIN
       TOBDETR:=TOB.Create('ECRITURE',TOBR,-1) ; TOBDETR.Dupliquer(TOBDest.Detail[i-1],True,True,False) ;
       TOBDETR.AddChampSup('INDICE',False) ; TOBDETR.PutValue('INDICE',i) ;
       TOBDETR.PutValue('E_DEBIT',DP) ; TOBDETR.PutValue('E_CREDIT',CP) ;
       TOBDETR.PutValue('E_DEBITEURO',DE) ; TOBDETR.PutValue('E_CREDITEURO',CE) ;
       TOBDETR.PutValue('E_DEBITDEV',DD) ; TOBDETR.PutValue('E_CREDITDEV',CD) ;
       TripoteDetail(TOBDETR) ;
       DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ; DE:=0 ; CE:=0 ;
       END ;
    XD:=GetLeMontant(TOBL.GetValue('E_DEBIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot) ;
    XC:=GetLeMontant(TOBL.GetValue('E_CREDIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot) ;
    DP:=Arrondi(DP+XD,V_PGI.OkDecV) ; CP:=Arrondi(CP+XC,V_PGI.OkDecV) ;

    XD:=GetLeMontant(TOBL.GetValue('E_DEBITEURO'),TOBL.GetValue('E_COUVERTUREEURO'),TOBL.GetValue('E_SAISIMP'),tsmEuro) ;
    XC:=GetLeMontant(TOBL.GetValue('E_CREDITEURO'),TOBL.GetValue('E_COUVERTUREEURO'),TOBL.GetValue('E_SAISIMP'),tsmEuro) ;
    DE:=Arrondi(DE+XD,V_PGI.OkDecE) ; CE:=Arrondi(CE+XC,V_PGI.OkDecE) ;

    XD:=GetLeMontant(TOBL.GetValue('E_DEBITDEV'),TOBL.GetValue('E_COUVERTUREDEV'),TOBL.GetValue('E_SAISIMP'),tsmDevise) ;
    XC:=GetLeMontant(TOBL.GetValue('E_CREDITDEV'),TOBL.GetValue('E_COUVERTUREDEV'),TOBL.GetValue('E_SAISIMP'),tsmDevise) ;
    DD:=Arrondi(DD+XD,ParmsMP.DEV.Decimale) ; CD:=Arrondi(CD+XC,ParmsMP.DEV.Decimale) ;
    OldAux:=TOBL.GetValue('E_AUXILIAIRE') ;
    OldGen:=TOBL.GetValue('E_GENERAL') ;
    OldEche:=TOBL.GetValue('E_DATEECHEANCE') ;
    END ;
if Not Premier then
   BEGIN
   TOBDETR:=TOB.Create('ECRITURE',TOBR,-1) ; TOBDETR.Dupliquer(TOBDest.Detail[Lasti],True,True,False) ;
   TOBDETR.AddChampSup('INDICE',False) ; TOBDETR.PutValue('INDICE',-1) ;
   TOBDETR.PutValue('E_DEBIT',DP) ; TOBDETR.PutValue('E_CREDIT',CP) ;
   TOBDETR.PutValue('E_DEBITEURO',DE) ; TOBDETR.PutValue('E_CREDITEURO',CE) ;
   TOBDETR.PutValue('E_DEBITDEV',DD) ; TOBDETR.PutValue('E_CREDITDEV',CD) ;
   TripoteDetail(TOBDETR) ;
   END ;
FinirDestination(TOBR) ;
TOBR.Free ;
END ;

procedure TFMulMPPRE.BOuvrirClick(Sender: TObject);
begin
  inherited;
ConstitueOrigines ;
if TOBOrig.Detail.Count<=0 then Exit ;
ParmsMP.DEV.Code:=E_DEVISE.Value ; GetInfosDevise(ParmsMP.DEV) ;
ConstitueDest ;
TOBDest.InsertDB(Nil) ;
end;

procedure TFMulMPPRE.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
TOBOrig.Free ;
TOBDest.Free ;
end;

end.

