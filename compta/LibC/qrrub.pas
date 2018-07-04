unit QRRub;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Hctrls, StdCtrls, ComCtrls, HRichEdt, HRichOLE, HQuickrp, hmsgbox, Menus,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DB, Buttons, ExtCtrls, Hcompte, Grids, HEnt1, Ent1,
  UObjFiltres {SG6 04/01/05 Gestion Filtres V6 FQ 15145},
  HSysMenu, UtilEdt, HTB97, HPanel, UiUtil, CpteUtil ;

procedure ImpRubrique(UnCpte : String ; DuBouton,Budget : boolean) ;

type
  TQRRubrique = class(TForm)
    Pages: TPageControl;
    Standards: TTabSheet;
    TFCpte: THLabel;
    HLabel1: THLabel;
    TFSigne: THLabel;
    TFaCpte: TLabel;
    FLibelle: TEdit;
    FSigne: THValComboBox;
    Option: TTabSheet;
    FApercu: TCheckBox;
    FTri: TRadioGroup;
    OptionsType: TGroupBox;
    DetailRub: TCheckBox;
    RapCrit: TCheckBox;
    FCouleur: TCheckBox;
    QRP: TQuickReport;
    SRub: TDataSource;
    QRub: TQuery;
    MsgRien: THMsgBox;
    TFType: THLabel;
    FType: THValComboBox;
    RB_FAMILLES: THMultiValComboBox;
    TFFamille: THLabel;
    Panel1: TPanel;
    BDTitreBis: TQRBand;
    QRShape20: TQRShape;
    TTitre2: TQRSysData;
    BDTitre: TQRBand;
    TTitre: TQRSysData;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel5: TQRLabel;
    TRType: TQRLabel;
    RRub1: TQRLabel;
    RRub2: TQRLabel;
    RLibelle: TQRLabel;
    RType: TQRLabel;
    BEntetePage: TQRBand;
    TitreEntete: TQRSysData;
    TitreBarre: TQRShape;
    piedpage: TQRBand;
    RCopyright: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    RSociete: TQRLabel;
    RUtilisateur: TQRLabel;
    BDFiche: TQRBand;
    BDFicheDetail: TQRBand;
    QRLabel3: TQRLabel;
    RSigne: TQRLabel;
    TRFamille: TQRLabel;
    RFamille: TQRLabel;
    Code: TQRLabel;
    RB_RUBRIQUE: TQRDBText;
    QRLabel6: TQRLabel;
    RB_LIBELLE: TQRDBText;
    QRLabel8: TQRLabel;
    RB_TYPERUB: TQRDBText;
    QRLabel9: TQRLabel;
    RB_SIGNERUB: TQRDBText;
    TRB_AXE: TQRLabel;
    RB_AXE: TQRDBText;
    FCompte1: TQRLabel;
    FExcept1: TQRLabel;
    FCompte2: TQRLabel;
    FExcept2: TQRLabel;
    QRShape1: TQRShape;
    QRLabel12: TQRLabel;
    FFamilles: TQRRichEdit;
    FFam: THRichEdit;
    BDEspace: TQRBand;
    FCpte1: THRichEdit;
    FEx1: THRichEdit;
    FCpte2: THRichEdit;
    FEx2: THRichEdit;
    FCalc: THRichEdit;
    FComptes1: TQRRichEdit;
    FExceptions1: TQRRichEdit;
    FComptes2: TQRRichEdit;
    FExceptions2: TQRRichEdit;
    FCalcul: TQRRichEdit;
    CbCaption: TComboBox;
    BDCol: TQRBand;
    Hor0: TQRShape;
    TLG_LIBELLE: TQRLabel;
    TLG_NATUREGENE: TQRLabel;
    TLG_SENS: TQRLabel;
    TLG_DATEMODIFICATION: TQRLabel;
    TLG_GENERAL: TQRLabel;
    QRShape17: TQRShape;
    BDLegende: TQRBand;
    TRLegende: TQRLabel;
    TLRSens: TQRLabel;
    TLRNat: TQRLabel;
    LRSens: TQRLabel;
    LRNat3: TQRLabel;
    LRSensBud: TQRLabel;
    BDListe: TQRBand;
    LRB_RUBRIQUE: TQRDBText;
    LRB_LIBELLE: TQRDBText;
    LRB_TYPERUB: TQRDBText;
    LRB_SIGNERUB: TQRDBText;
    LRB_AXE: TQRDBText;
    LRB_FAMILLES: TQRDBText;
    TypeEdition: TRadioGroup;
    Colonnes: TCheckBox;
    Overlay: TQRBand;
    Col0: TQRLigne;
    Col3: TQRLigne;
    Col1: TQRLigne;
    Col2: TQRLigne;
    Col4: TQRLigne;
    Col5: TQRLigne;
    Col6: TQRLigne;
    Hor2: TQRLigne;
    Hor1: TQRLigne;
    BDFoot: TQRBand;
    BDSummary: TQRBand;
    QRubRB_RUBRIQUE: TStringField;
    QRubRB_LIBELLE: TStringField;
    QRubRB_FAMILLES: TStringField;
    QRubRB_SIGNERUB: TStringField;
    QRubRB_TYPERUB: TStringField;
    QRubRB_COMPTE1: TStringField;
    QRubRB_COMPTE2: TStringField;
    QRubRB_AXE: TStringField;
    RNumversion: TQRLabel;
    HMTrad: THSystemMenu;
    FRub1: THValComboBox;
    FRub2: THValComboBox;
    TLG_FAMILLES: TQRLabel;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    QRubRB_EXCLUSION1: TStringField;
    QRubRB_EXCLUSION2: TStringField;
    Dock971: TDock97;
    PanelFiltre: TToolWindow97;
    HPB: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BDTitreAfterPrint(BandPrinted: Boolean);
    procedure BFermeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure piedpageAfterPrint(BandPrinted: Boolean);
    procedure BDTitreBisAfterPrint(BandPrinted: Boolean);
    procedure BDFicheBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure SRubDataChange(Sender: TObject; Field: TField);
    procedure BDFicheDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TypeEditionClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure RB_FAMILLESChange(Sender: TObject);
    procedure BDLegendeBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private    { Déclarations privées }
    //SG6 04/01/05 Gestion Filtres V6 FQ 15145
    ObjFiltre : TObjFiltre;
    LeCpte                  : string ;
    Budget,QuiMappel,Liste,Fiche   : boolean ;
    FAM,TYP,CP1,EX1,CP2,EX2 : TField ;
    Filtre : string;
    procedure ChoixEdition ;
    procedure GenereSQL ;
    procedure RenseigneCritere ;
    procedure IndexRub ;
    procedure RemplitMemo(Avec:TField; Memo,MemoCalc:THRichEdit; QuelCas:integer; TT:String) ;
    procedure RemplitDetail(TypeRub:String);
    procedure InitDivers ;
  public
    { Déclarations publiques }
  end;


implementation

//XMG 05/04/04 début
uses
  {$IFDEF ESP}
  //FAMRUB_TOF
  Rubrique_TOM
  {$ELSE}
  FamRub
  {$ENDIF ESP}
  ;
  //XMG 05/04/04 fin

{$R *.DFM}

procedure ImpRubrique(UnCpte : String ; DuBouton,Budget : Boolean) ;
var QRRubrique: TQRRubrique;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QRRubrique:=TQRRubrique.Create(Application) ;
QRRubrique.QuiMappel:=DuBouton ;
QRRubrique.Budget:=Budget ;
if Budget then QRRubrique.Filtre:='IMPRUBBUD'
          else QRRubrique.Filtre:='IMPRUB' ;
if PP=Nil then
   BEGIN
   try
    if Not DuBouton then QRRubrique.ShowModal
                    else BEGIN QRRubrique.LeCpte:=UnCpte ; QRRubrique.BValiderClick(Nil) ; END ;
    finally
    QRRubrique.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(QRRubrique,PP) ;
   QRRubrique.Show ;
   END ;
END ;

procedure TQRRubrique.FormShow(Sender: TObject);
Var i : Integer ;
begin
if Budget then Caption:=MsgRien.Mess[4] ;
QRP.ReportTitle:=Caption ;
if Budget then
  BEGIN
  LRB_TYPERUB.DataField:=LRB_FAMILLES.DataField ; LRB_FAMILLES.DataField:='' ;
  TFType.Visible:=False ; FType.Visible:=False ;
  TFSigne.Top:=TFType.Top ; FSigne.Top:=FType.Top ;
  TFFamille.Caption:=TFType.Caption ;
  RB_FAMILLES.DataType:='' ; FType.DataType:='' ;
  RempliComboFamRub(FType) ;
  RB_FAMILLES.Items.Clear ; RB_FAMILLES.Values.Clear ;
  for i:=0 to FType.Items.Count-1 do
    BEGIN
    RB_FAMILLES.Items.Add(FType.Items[i]) ;
    RB_FAMILLES.Values.Add(FType.Values[i]) ;
    END ;
  //XMG 05/04/04 début
  {$IFDEF ESP}
  FactoriseComboTypeRub(FType) ;
  {$ELSE}
  RempliComboTypRub(FType) ;
  {$ENDIF ESP}

  END ;
FSigne.ItemIndex:=0 ; FType.ItemIndex:=0 ; //RB_FAMILLES.Tous:=True ;
Pages.ActivePage:=Standards ;
FCouleur.Checked:=V_PGI.QRCouleur ;
RB_FAMILLESChange(Nil) ;

//Sg6 04/01/05 Gestion Filtres V6
ObjFiltre.FFI_TABLE:=Filtre;
ObjFiltre.Charger;

end;

procedure TQRRubrique.BValiderClick(Sender: TObject);
begin
SourisSablier ;
EnableControls(Self,False) ;
if QuiMappel then
   BEGIN
   RB_FAMILLESChange(nil) ;
   FRub1.Value:=LeCpte ; FRub2.Value:=LeCpte ; {TypeEdition.ItemIndex:=1} ;
   BDTitreBis.Enabled:=False ;
   if Budget then QRP.ReportTitle:=MsgRien.Mess[5]
             else QRP.ReportTitle:=MsgRien.Mess[1] ;  //CRi 22/05/97
   END ;
RedimLargeur(Self, QRP,FCouleur.checked, TTitre);
ChoixEdition ;
GenereSQL ;
if Not QuiMappel then RenseigneCritere ;
InitDivers ;
QRub.Close ; ChangeSQL(QRub) ; QRub.Open ; IndexRub ;
if Budget then BEGIN RemplitMemo(FAM,FFam,Nil,0,'ttRubTypeBud') ; END
          else RemplitMemo(FAM,FFam,Nil,0,'ttRubFamille') ;
if DetailRub.Checked then
   BEGIN
   RemplitMemo(CP1,FCpte1,FCalc,2,'ttRubCalcul') ;
   RemplitMemo(CP2,FCpte2,Nil,2,'') ;
   RemplitMemo(EX1,FEx1,Nil,1,'') ;
   RemplitMemo(EX2,FEx2,Nil,1,'') ;
   END ;
if QuiMappel then QRP.QRPrinter.SynCouleur:=V_PGI.QRCouleur else QRP.QRPrinter.SynCouleur:=FCouleur.Checked ;
// if QuiMappel then BEGIN BDTitre.Enabled:=False ; QRP.ReportTitle:=MsgRien.Mess[1] ; END ;    //  CRi 22/05/97
if QRub.Eof then MsgRien.Execute(0,Caption,'')
             else BEGIN if FApercu.Checked then QRP.Preview else QRP.Print ; END ;
QRub.Close ;
EnableControls(Self,True) ;
SourisNormale ;
end;

procedure TQRRubrique.ChoixEdition ;
var i : integer ;
BEGIN
if QuiMappel then TypeEdition.ItemIndex:=1 ;
Liste:=(TypeEdition.ItemIndex=0) ; Fiche:=(TypeEdition.ItemIndex=1) ;
BDListe.Enabled:=Liste ; BDCol.Enabled :=Liste ; BDEspace.Enabled:=Fiche ;
BDFiche.Enabled:=Fiche ; BDFicheDetail.Enabled:=(Fiche and DetailRub.Checked) ;
BDFiche.Frame.DrawBottom:=(Not DetailRub.Checked) ;
BDTitre.Enabled:=(RapCrit.Checked) ;
if Not QuiMappel then BDTitreBis.Enabled:=Not (RapCrit.Checked) ;  //  CRi 22/05/97
BDLegende.Enabled:=(Liste and (RapCrit.Checked)) ;
Hor1.Visible:=Liste ; Hor2.Visible:=Liste ;
Col0.Visible:=Liste ; Col6.Visible:=Liste ;
{ Colonnages }
For i:=1 to 5 do TQRLigne(FindComponent('Col'+IntToStr(i))).Visible:=Colonnes.Checked ;
END ;

procedure TQRRubrique.GenereSQL ;
Var Traduction,St,StOr : string ;
BEGIN
QRub.SQL.Clear ;
QRub.SQL.Add('Select * From RUBRIQUE ') ;
QRub.SQL.Add('WHERE RB_RUBRIQUE<>"'+W_W+'" ') ;
if Budget then QRub.SQL.Add(' AND RB_NATRUB="BUD" AND RB_BUDJAL<>"" ')
          Else QRub.SQL.Add(' AND RB_NATRUB<>"BUD" AND RB_BUDJAL="" ') ;
if NOT QuiMappel then
   BEGIN
    if FRub1.Value<>'' then QRub.SQL.Add(' And RB_RUBRIQUE>="'+FRub1.Value+'"') ;
    if FRub2.Value<>'' then QRub.SQL.Add(' And RB_RUBRIQUE<="'+FRub2.Value+'"') ;
    if FLibelle.Text<>'' then
       BEGIN
       Traduction:=TraduitJoker(FLibelle.text) ;
       QRub.SQL.Add(' and UPPER(RB_LIBELLE) like "'+Traduction+'" ' );
       END ;
    if FSigne.ItemIndex<>0 then QRub.SQL.Add(  ' And RB_SIGNERUB="'+FSigne.Value+'"') ;
    if FType.ItemIndex<>0 then QRub.SQL.Add(  ' And RB_TYPERUB="'+FType.Value+'"') ;
    if Not RB_FAMILLES.Tous then
       BEGIN
       St:=RB_FAMILLES.Text ;
       if St[Length(St)]<>';' then St:=St+';' ;
       StOr:='' ;
       While St<>'' do StOr:=StOr+' OR '+RB_FAMILLES.Name+' like "%'+ReadTokenSt(St)+'%"' ;
       System.Delete(StOr,1,4) ;
       if StOR<>'' then QRub.SQL.Add(' AND ('+StOr+')') ;
       END ;
    Case FTri.ItemIndex of
       0 : QRub.SQL.Add(' Order By RB_RUBRIQUE') ;
       1 : QRub.SQL.Add(' Order By RB_LIBELLE') ;
     end ;
   END else
   BEGIN
   QRub.SQL.Add(' And RB_RUBRIQUE>="'+FRub1.Value+'"') ; QRub.SQL.Add(' And RB_RUBRIQUE<="'+FRub2.Value+'"') ;
   END ;
END;

procedure TQRRubrique.RenseigneCritere ;
BEGIN
RRub1.Caption:=FRub1.Text ; RRub2.Caption:=FRub2.Text ; RLibelle.Caption:=FLibelle.Text ;
RSigne.Caption:=FSigne.Text ; RType.Caption:=FType.Text ;
if Budget then
  BEGIN
  TRType.Visible:=False ; RType.Visible:=False ;
  TRFamille.Caption:=TFFamille.Caption ;
  END ;
RFamille.Caption:=RB_FAMILLES.Text ;
END ;

procedure TQRRubrique.InitDivers ;
BEGIN
RSociete.Caption:=Copy(RSociete.Caption,1,Pos(':',RSociete.Caption))
                       +' '+V_PGI.CodeSociete+' '+V_PGI.NomSociete ;
RUtilisateur.Caption:=Copy(RUtilisateur.Caption,1,Pos(':',RUtilisateur.Caption))
                       +' '+V_PGI.User+' '+V_PGI.UserName ;
RNumversion.Caption:=MsgRien.Mess[2]+' '+V_PGI.NumVersion+' '+MsgRien.Mess[3]+' '+DateToStr(V_PGI.DateVersion);
RCopyright.Caption := Copyright + ' - ' + TitreHalley ;
END ;

procedure TQRRubrique.IndexRub ;
BEGIN
FAM:=QRub.FindField('RB_FAMILLES') ; TYP:=QRub.FindField('RB_TYPERUB') ;
CP1:=QRub.FindField('RB_COMPTE1') ;  EX1:=QRub.FindField('RB_EXCLUSION1') ;
CP2:=QRub.FindField('RB_COMPTE2') ;  EX2:=QRub.FindField('RB_EXCLUSION2') ;
END ;


procedure TQRRubrique.BDTitreAfterPrint(BandPrinted: Boolean);
begin BEntetePage.Enabled:=False ; end;

procedure TQRRubrique.BFermeClick(Sender: TObject);
begin
Close ;
//SG6 04/01/05 Vide le panel
if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TQRRubrique.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
//SG6 04/01/05 Gestion Filtres V6 FQ 15145
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants,'');

PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
HorzScrollBar.Range:=0 ; HorzScrollBar.Visible:=FALSE ;
VertScrollBar.Range:=0 ; VertScrollBar.Visible:=FALSE ;
Pages.Top:=0 ; Pages.Left:=0 ;
ClientHeight:=Pages.Top+Pages.Height+HPB.Height+PanelFiltre.Height-1 ;
ClientWidth:=Pages.Left+Pages.Width ;
if V_PGI.OutLook then Pages.Align:=alClient ;
end;

procedure TQRRubrique.piedpageAfterPrint(BandPrinted: Boolean);
begin BEntetePage.Enabled:=True ; end;

procedure TQRRubrique.BDTitreBisAfterPrint(BandPrinted: Boolean);
begin BEntetePage.Enabled:=False ; end ;

procedure TQRRubrique.BDFicheBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
if RB_TYPERUB.Caption<>'' then RB_TYPERUB.Caption:=RechDom('ttRubType',RB_TYPERUB.Caption,False) ;
if RB_SIGNERUB.Caption<>'' then RB_SIGNERUB.Caption:=RechDom('ttRubSigne',RB_SIGNERUB.Caption,False) ;
TRB_AXE.Visible:=(Pos('A',TYP.AsString)>0) ; RB_AXE.Visible:=(Pos('A',TYP.AsString)>0) ;
if RB_AXE.Caption<>'' then RB_AXE.Caption:=RechDom('ttAxe',RB_AXE.Caption,False) ;
end;

function TransBud(St : String) : String ;
var s : String ;
BEGIN
s:='' ;
if St='CBG'then s:='GEN' else
 if St='CBS' then s:='ANA' else
  if St='G/S' then s:='G/A' else
   if St='S/G' then s:='A/G' ;
Result:=s ;
END ;

procedure TQRRubrique.RemplitMemo(Avec:TField; Memo,MemoCalc:THRichEdit; QuelCas:integer; TT:String) ;
var St,STC,StB,STL : string ;
    j          : integer ;
begin
St:=Avec.AsString ; Memo.Lines.Clear ; if MemoCalc<>Nil then MemoCalc.Lines.Clear ;
While St<>'' do
   BEGIN
   STC:=ReadTokenSt(St) ; if STC='' then Break ;
   case QuelCas of
      0 : BEGIN                                    //affichage des familles
          STB:=STC ;
          if Budget then STB:=TransBud(StC) ;
          if TT<>'' then STL:=RechDom(TT,STB,False) ;
          Memo.Lines.Add('('+STC+') - '+ STL) ;
          END ;
      1 : Memo.Lines.Add(STC) ;                    //affichage des comptes exclus
      2 : BEGIN                                    //affichage des comptes et des modes de calcul
          if Pos('(',STC)>0 then
             BEGIN
             j:=Pos('(',STC) ; STL:=Copy(STC,j+1,2) ;
             if TT<>'' then STL:=RechDom(TT,STL,False) ;
             if MemoCalc<>Nil then MemoCalc.Lines.Add(STL) ;
             STC:=Copy(STC,1,j-1) ; Memo.Lines.Add(STC) ;
             END ;
          END ;
   end ;
   END ;
end;

procedure TQRRubrique.SRubDataChange(Sender: TObject; Field: TField);
begin
if Budget then RemplitMemo(QRub.FindField('RB_FAMILLES'),FFam,Nil,0,'ttRubTypeBud')
          else RemplitMemo(QRub.FindField('RB_FAMILLES'),FFam,Nil,0,'ttRubFamille') ;
if DetailRub.Checked then
   BEGIN
   RemplitMemo(QRub.FindField('RB_COMPTE1'),FCpte1,FCalc,2,'ttRubCalcul') ;
   RemplitMemo(QRub.FindField('RB_COMPTE2'),FCpte2,Nil,2,'') ;
   RemplitMemo(QRub.FindField('RB_EXCLUSION1'),FEx1,Nil,1,'') ;
   RemplitMemo(QRub.FindField('RB_EXCLUSION2'),FEx2,Nil,1,'') ;
   END ;
end;

procedure TQRRubrique.RemplitDetail(TypeRub:String);
Var C1,C2 : Char ;
    Double : boolean ;
BEGIN
C1:=TypeRub[1] ; C2:=TypeRub[3] ; Double:=(TypeRub[2]='/') ;
Case C1 of
 'A':BEGIN FCompte1.Caption:=CbCaption.Items[0] ; FExcept1.Caption:=CbCaption.Items[1] ; END ;
 'B':BEGIN FCompte1.Caption:=CbCaption.Items[2] ; FExcept1.Caption:=CbCaption.Items[3] ; END ;
 'T':BEGIN FCompte1.Caption:=CbCaption.Items[4] ; FExcept1.Caption:=CbCaption.Items[5] ; END ;
 'G':BEGIN FCompte1.Caption:=CbCaption.Items[6] ; FExcept1.Caption:=CbCaption.Items[7] ; END ;
 End ;
FCompte2.Visible:=Double ; FExcept2.Visible:=Double ;
FComptes2.Visible:=Double ; FExceptions2.Visible:=Double ;
if Double then
   BEGIN
   FCompte1.Width:=150 ; FExcept1.Left:=160 ; FExcept1.Width:=120 ;
   FComptes1.Width:=150 ; FExceptions1.Left:=160 ; FExceptions1.Width:=120 ;
    Case C2 of
      'A':BEGIN FCompte2.Caption:=CbCaption.Items[0] ; FExcept2.Caption:=CbCaption.Items[1] ; END ;
      'B':BEGIN FCompte2.Caption:=CbCaption.Items[2] ; FExcept2.Caption:=CbCaption.Items[3] ; END ;
      'T':BEGIN FCompte2.Caption:=CbCaption.Items[4] ; FExcept2.Caption:=CbCaption.Items[5] ; END ;
      'G':BEGIN FCompte2.Caption:=CbCaption.Items[6] ; FExcept2.Caption:=CbCaption.Items[7] ; END ;
      End ;
   END else
   BEGIN
   FCompte1.Width:=300 ; FExcept1.Left:=310 ; FExcept1.Width:=240 ;
   FComptes1.Width:=300 ; FExceptions1.Left:=310 ; FExceptions1.Width:=240 ;
   END ;
END ;

procedure TQRRubrique.BDFicheDetailBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
RemplitDetail(TYP.AsString) ;
end;

procedure TQRRubrique.TypeEditionClick(Sender: TObject);
begin
DetailRub.Enabled:=(TypeEdition.ItemIndex=1) ;
DetailRub.Checked:=(TypeEdition.ItemIndex=1) ;
Colonnes.Checked:=Not (TypeEdition.ItemIndex=1) ;
Colonnes.Enabled:=Not (TypeEdition.ItemIndex=1) ;
end;

procedure TQRRubrique.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TQRRubrique.RB_FAMILLESChange(Sender: TObject);
Var Q  : TQuery ;
    StSql, StFam, St, StOr : String ;
begin
StFam:='' ;
StSql:='Select RB_RUBRIQUE, RB_LIBELLE from RUBRIQUE Where RB_RUBRIQUE<>"'+W_W+'" ';
if Budget then StSql :=StSql+' AND RB_NATRUB="BUD" AND RB_BUDJAL<>"" '
          Else StSql :=StSql+' AND RB_NATRUB<>"BUD" AND RB_BUDJAL="" ' ;

if Not RB_FAMILLES.Tous then
   BEGIN
   St:=RB_FAMILLES.Text ;
   if St[Length(St)]<>';' then St:=St+';' ;
   StOr:='' ;
   While St<>'' do StOr:=StOr+' OR '+RB_FAMILLES.Name+' like "%'+ReadTokenSt(St)+'%"' ;
   System.Delete(StOr,1,4) ;
   if StOR<>'' then StFam:=StFam+' AND ('+StOr+')' ;
   END ;
StSql:=StSql+StFam+' ORDER BY RB_RUBRIQUE ' ;
Q:=OpenSQL(StSql,True) ;
FRub1.Items.Clear ; FRub2.Items.Clear ;
FRub1.Values.Clear ; FRub2.Values.Clear ;
While not Q.Eof do
      BEGIN
      FRub1.Values.Add(Q.Fields[0].AsString);
      FRub2.Values.Add(Q.Fields[0].AsString);
      FRub1.Items.Add(Q.Fields[1].AsString);
      FRub2.Items.Add(Q.Fields[1].AsString);
      Q.Next ;
      END ;
FRub1.ItemIndex:=0 ; FRub2.ItemIndex:=FRub2.Items.Count-1 ;
Ferme(Q) ;
end;

procedure TQRRubrique.BDLegendeBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
LRSens.Visible:=not Budget ; LRSensBud.Visible:=Budget ;
end;


procedure TQRRubrique.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//SG6 04/01/05 Gestion Filtres V6 FQ 15145
ObjFiltre.Free;
if Parent is THPanel then Action:=caFree ;
end;

end.
