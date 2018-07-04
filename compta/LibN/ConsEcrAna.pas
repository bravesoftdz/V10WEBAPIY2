unit ConsEcrAna ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB, DBTables, Hqry,
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, SaisUtil,
  HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil, UTOB,
  General, CritEDT, UtilEDT,

  QRGLGen, QRGLAna, QRCumulG, HRichOLE, FE_Main,
{$IFNDEF CCMP}
  Outils,
{$ENDIF}
  CumMens,
{$IFNDEF CCMP}
  MulAna,
{$ENDIF}
  Section,
  SaisODA ;

procedure OperationsSurSections ( LeGene,LeExo,ChampsTries : string; LaSect : string='' ; FromSaisie : boolean = False ) ;

type
  TFConsEcrAna = class(TFMul)
    HM: THMsgBox;
    Y_VALIDE: TCheckBox;
    TY_REFINTERNE: THLabel;
    Y_REFINTERNE: TEdit;
    Y_SAISIEEURO: TCheckBox;
    PEcritures: TTabSheet;
    Bevel6: TBevel;
    TY_JOURNAL: THLabel;
    TY_EXERCICE: THLabel;
    TY_NATUREPIECE: THLabel;
    TY_NUMEROPIECE: THLabel;
    HLabel1: THLabel;
    TY_DATECOMPTABLE: THLabel;
    TY_DATECOMPTABLE2: THLabel;
    Y_NUMEROPIECE: THCritMaskEdit;
    Y_NUMEROPIECE_: THCritMaskEdit;
    Y_JOURNAL: THValComboBox;
    Y_EXERCICE: THValComboBox;
    Y_NATUREPIECE: THValComboBox;
    Y_DATECOMPTABLE: THCritMaskEdit;
    Y_DATECOMPTABLE_: THCritMaskEdit;
    TY_GENERAL: THLabel;
    Y_GENERAL: THCpteEdit;
    TY_SECTION: THLabel;
    Y_SECTION: THCpteEdit;
    TGLIBELLE: THLabel;
    TSLIBELLE: THLabel;
    BActionCPT: TToolbarButton97;
    PopACPT: TPopupMenu;
    ModifGENE: TMenuItem;
    BActionEcr: TToolbarButton97;
    PopAEdt: TPopupMenu;
    ModifEcr: TMenuItem;
    XX_WHERE1: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    TSoldes: TLabel;
    PopZ: TPopupMenu;
    N5: TMenuItem;
    XX_WHEREAN: TEdit;
    XX_WHERELET: TEdit;
    PParams: TTabSheet;
    Bevel7: TBevel;
    TLeSolde: THLabel;
    RDefil: TRadioGroup;
    RAction: TRadioGroup;
    N6: TMenuItem;
    ROppose: TRadioGroup;
    TY_ETABLISSEMENT: THLabel;
    Y_ETABLISSEMENT: THValComboBox;
    GeneUp: TToolbarButton97;
    GeneDown: TToolbarButton97;
    SectUp: TToolbarButton97;
    SectDown: TToolbarButton97;
    Y_CREERPAR: THCritMaskEdit;
    CAvecSimu: TCheckBox;
    XX_WHEREQUALIF: TEdit;
    XX_WHEREDET: TEdit;
    XX_WHEREDC: TEdit;
    LESOLDE: THNumEdit;
    N8: TMenuItem;
    ModifSection: TMenuItem;
    GLAnalyt: TMenuItem;
    PzLibre: TTabSheet;
    Bevel5: TBevel;
    TY_TABLE0: THLabel;
    Y_TABLE0: THCpteEdit;
    Y_TABLE2: THCpteEdit;
    Y_TABLE1: THCpteEdit;
    Y_TABLE3: THCpteEdit;
    TY_TABLE2: THLabel;
    TY_TABLE1: THLabel;
    TY_TABLE3: THLabel;
    TY_DEVISE: THLabel;
    Y_DEVISE: THValComboBox;
    TY_AXE: THLabel;
    Y_AXE: THValComboBox;
    CumulSect: TMenuItem;
    CumulsGene: TMenuItem;
    GLGene: TMenuItem;
    SaisieODA: TMenuItem;
    N1: TMenuItem;
    Y_TYPEANALYTIQUE: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure Y_EXERCICEChange(Sender: TObject);
    procedure Y_GENERALExit(Sender: TObject);
    procedure Y_SECTIONExit(Sender: TObject);
    procedure Y_GENERALEnter(Sender: TObject);
    procedure Y_SECTIONEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ModifGENEClick(Sender: TObject);
    procedure CumulsGENEClick(Sender: TObject);
    procedure GLGENEClick(Sender: TObject);
    procedure ModifEcrClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject); Override ;
    procedure PopACPTPopup(Sender: TObject);
    procedure PopAEdtPopup(Sender: TObject);
    procedure PopZPopup(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure RActionClick(Sender: TObject);
    procedure ROpposeClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure SectUpClick(Sender: TObject);
    procedure SectDownClick(Sender: TObject);
    procedure GeneUpClick(Sender: TObject);
    procedure GeneDownClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure CAvecSimuClick(Sender: TObject);
    procedure GLAnalytClick(Sender: TObject);
    procedure ModifSectionClick(Sender: TObject);
    procedure CumulSectClick(Sender: TObject);
    procedure SaisieODAClick(Sender: TObject);
    procedure Y_AXEChange(Sender: TObject);
  private
    OldGene,OldSect : String ;
    ToutAcc,GeneCharge,FocusUD : Boolean ;
    TOBGene,TOBSection : TOB ;
    DroitEcritures : boolean ;
    procedure InitCriteres ;
    procedure AttribTitre ;
    Function  PositionneExo : TExoDate ;
    procedure EnabledMenus ;
    procedure GotoEntete ;
    procedure AfficheLeSoldeUnique ;
    procedure GeneSuivant ( Suiv : boolean ) ;
    procedure SectSuivant ( Suiv : boolean ) ;
    procedure CacheZones ;
    procedure GereInfoOrig ;
    procedure ChargeGeneral ( Force : boolean ) ;
    procedure ChargeSection ( Force : boolean ) ;
    procedure EtudieANouveau ;
    procedure EtudieConvert ;
    Function  RecupWhereDataType : String ;
  public
    LeGene,LeExo,ChampsTries,LaSect : String ;
    FromSaisie                     : boolean ;
  end;

implementation

{$R *.DFM}

procedure OperationsSurSections ( LeGene,LeExo,ChampsTries : string; LaSect : string='' ; FromSaisie : boolean = False ) ;
var X : TFConsEcrAna ;
    M : RMVT ;
    PP : THPanel ;
begin
X:=TFConsEcrAna.Create(Application) ;
X.TypeAction:=taModif ; X.FNomFiltre:='CPOPERSECT' ;
X.Q.Manuel:=True ; X.Q.Liste:='CPCONSULTSECT' ;
X.LeGene:=LeGene ; X.LeExo:=LeExo ; X.ChampsTries:=ChampsTries ;
X.LaSect:=LaSect ; X.FromSaisie:=FromSaisie ;
PP:=FindInsidePanel ;
if ((PP=Nil) or (V_PGI.ZoomOLE)) then
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

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 21/11/2001
Modifié le ... :   /  /    
Description .. : Modification de la section de la ligne en cours.
Mots clefs ... : 
*****************************************************************}
procedure TFConsEcrAna.ModifSectionClick(Sender: TObject);
var  AA  : TActionFiche ;
begin
  inherited;
 if (Q.Eof) And (Q.Bof)  then Exit ;
 AA:=taModif ; if Not ToutAcc then AA:=taConsult ;
 FicheSection(Nil,Q.FindField('Y_AXE').AsString,Q.FindField('Y_SECTION').AsString,AA,0) ;
end;

procedure TFConsEcrAna.AttribTitre ;
BEGIN
//HelpContext:=0 ;
UpdateCaption(Self) ;
if VH^.TenueEuro then Y_SAISIEEURO.Caption:=HM.Mess[8]+' '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False) ;
END ;

procedure TFConsEcrAna.EtudieANouveau ;
Var DD : TDateTime ;
BEGIN
XX_WHEREAN.Text:='Y_ECRANOUVEAU="N"' ;
DD:=StrToDate(Y_DATECOMPTABLE.Text) ; if (VH^.Suivant.Deb>0) And (DD>VH^.Suivant.Deb) then Exit ;
XX_WHEREAN.Text:='Y_ECRANOUVEAU="N" OR ((Y_ECRANOUVEAU="H" OR Y_ECRANOUVEAU="OAN") AND Y_DATECOMPTABLE="'+UsDateTime(DD)+'")' ;
END ;

procedure TFConsEcrAna.EtudieConvert ;
BEGIN
if ((Y_DEVISE.Value=V_PGI.DevisePivot) or (Y_DEVISE.Value='')) then
   BEGIN
   if ROppose.ItemIndex=0 then
      BEGIN
      XX_WHEREDC.Text:='Y_DEBIT<>0 OR Y_CREDIT<>0' ;
      END else
      BEGIN
      XX_WHEREDC.Text:='Y_DEBITEURO<>0 OR Y_CREDITEURO<>0' ;
      END ;
   END else
   BEGIN
   XX_WHEREDC.Text:='' ;
   END ;
END ;

procedure TFConsEcrAna.CacheZones ;
Var i,idp,icp,ide,ice : integer ;
    VisuGene,VisuSect : boolean ;
BEGIN
if (ctxPCL in V_PGI.PGIContexte)=FALSE Then Exit ;
VisuGene:=True ; VisuSect:=True ;
idp:=-1 ; icp:=-1 ; ide:=-1 ; ice:=-1 ;
if TOBGene.GetValue('G_GENERAL')<>'' then
   BEGIN
   VisuGene:=False ;
   if TOBGene.GetValue('G_VENTILABLE')='X' then
      BEGIN
      if TOBSection.GetValue('S_SECTION')<>'' then VisuSect:=False ;
      END else
      BEGIN
      VisuSect:=False ;
      END ;
   END else
   BEGIN
   VisuSect:=False ;
   END ;
for i:=0 to FListe.Columns.Count-1 do
    BEGIN
    if FListe.Columns[i].FieldName='Y_GENERAL' then FListe.Columns[i].Visible:=VisuGene ;
    if FListe.Columns[i].FieldName='Y_SECTION' then FListe.Columns[i].Visible:=VisuSect ;
    if FListe.Columns[i].FieldName='Y_DEBIT' then idp:=i ;
    if FListe.Columns[i].FieldName='Y_CREDIT' then icp:=i ;
    if FListe.Columns[i].FieldName='Y_DEBITEURO' then ide:=i ;
    if FListe.Columns[i].FieldName='Y_CREDITEURO' then ice:=i ;
    END ;
if ROppose.itemIndex=0 then
   BEGIN
   if idp>0 then FListe.Columns[idp].Visible:=True ;
   if icp>0 then FListe.Columns[icp].Visible:=True ;
   if ide>0 then FListe.Columns[ide].Visible:=False ;
   if ice>0 then FListe.Columns[ice].Visible:=False ;
   END else
   BEGIN
   if ide>0 then FListe.Columns[ide].Visible:=True ;
   if ice>0 then FListe.Columns[ice].Visible:=True ;
   if idp>0 then FListe.Columns[idp].Visible:=False ;
   if icp>0 then FListe.Columns[icp].Visible:=False ;
   END ;
if ((HMTrad.ActiveResize) and (HMTrad.ResizeDBGrid)) then
   BEGIN
   HMTrad.ResizeDBGridColumns(FListe) ;
   FListe.Refresh ;
   END ;
END ;

procedure TFConsEcrAna.BChercheClick(Sender: TObject);
Var SavePlace: TBookmark;
begin
EtudieANouveau ;
Y_GENERALExit(Nil) ; Y_SECTIONExit(Nil) ;
if ((TOBGene.GetValue('G_GENERAL')='') and (TOBSection.GetValue('S_SECTION')='')) then
   if Not GeneCharge then
   BEGIN
   HM.Execute(9,Caption,'') ;
   if Y_GENERAL.CanFocus then Y_GENERAL.SetFocus ;
   Exit ;
   END ;
EtudieConvert ;
if (Sender=Nil) and Not((Q.EOF) and (Q.BOF)) then SavePlace:=Q.GetBookmark else SavePlace:=Nil ;
  inherited;
CacheZones ;
if (Sender=Nil) and Not((Q.EOF) and (Q.BOF)) then if SavePlace<>Nil then
   BEGIN
   Try
   Q.GotoBookmark(SavePlace) ;
   Q.FreeBookmark(SavePlace) ;
   Except
   End ;
   END ;
end;

procedure TFConsEcrAna.GereInfoOrig ;
BEGIN
if LeGene<>'' then
   BEGIN
   Y_GENERAL.Text:=LeGene ; ChargeGeneral(False) ;
   if LaSect<>'' then Y_SECTION.Text:=LaSect else Y_SECTION.Text:='' ;
   ChargeSection(False) ;
   XX_WHERE1.Text:='' ;
   END ;
if LeExo='0' then Y_EXERCICE.Value:=VH^.Encours.Code else
 if LeExo='1' then Y_EXERCICE.Value:=VH^.Suivant.Code else
  if LeExo='-1' then Y_EXERCICE.Value:=VH^.Precedent.Code ;
END ;

procedure TFConsEcrAna.FormShow(Sender: TObject);
begin
MakeZoomOLE(Handle) ;
GeneCharge:=True ;
DroitEcritures:=ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False) ;
if Not DroitEcritures then BEGIN RAction.ItemIndex:=0 ; RAction.Enabled:=False ; END ;
InitCriteres ; LibellesTableLibre(PzLibre,'TY_TABLE','Y_TABLE','A') ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   Y_EXERCICE.Value:=VH^.CPExoRef.Code ;
   Y_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   Y_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   Y_EXERCICE.Value:=VH^.Entree.Code ;
   Y_DATECOMPTABLE.Text:=DateToStr(VH^.Entree.Deb) ;
   Y_DATECOMPTABLE_.Text:=DateToStr(VH^.Entree.Fin) ;
   END ;
PositionneEtabUser(Y_ETABLISSEMENT) ;
AttribTitre ;
XX_WHEREAN.Text:='Y_ECRANOUVEAU="N"' ;
TGLibelle.Caption:='' ; TSLibelle.Caption:='' ; OldGene:='' ; OldSect:='' ;
ChangeMask(LeSolde,V_PGI.OkDecV,'') ;
GereInfoOrig ;
if ctxPCL in V_PGI.PGIContexte then CAvecSimu.Visible:=False ;
  inherited;
SetTousCombo(Y_JOURNAL) ; SetTousCombo(Y_DEVISE) ; SetTousCombo(Y_NATUREPIECE) ;
//SetTousCombo(Y_AXE) ;
Y_AXE.ItemIndex := 0 ;
if ctxPCL in V_PGI.PGIContexte then BParamListe.Visible:=False ;
GeneCharge:=False ;
//PSQL.TabVisible:=False ;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
XX_WHERE1.Text:='' ;
CacheZones ;
end;

procedure TFConsEcrAna.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then Y_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else Y_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
Y_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 26/11/2001
Modifié le ... :   /  /    
Description .. : Modification de l'écriture analytique sur double clic dans la 
Suite ........ : liste.
Mots clefs ... : 
*****************************************************************}
procedure TFConsEcrAna.FListeDblClick(Sender: TObject);
Var AA  : TActionFiche ;
begin
  inherited;
if (Q.Eof) And (Q.Bof) then Exit ;
AA:=taModif ; if Not ToutAcc then AA:=taConsult ;
TrouveEtLanceSaisieODA(Q,AA) ;
ChargeGeneral(True) ; ChargeSection(True) ;
BChercheClick(Nil) ;
end;

procedure TFConsEcrAna.Y_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(Y_EXERCICE.Value,Y_DATECOMPTABLE,Y_DATECOMPTABLE_) ;
if Y_EXERCICE.Value='' then BEGIN Y_DATECOMPTABLE.Text:=stDate1900 ; Y_DATECOMPTABLE_.Text:=stDate2099 ; END ;
end;

procedure TFConsEcrAna.ChargeGeneral ( Force : boolean ) ;
var tz : tzoomtable ;
BEGIN
if ((TOBGene=Nil) or (csDestroying in ComponentState)) then Exit ;
AfficheLeSoldeUnique ;
  inherited;
if Y_GENERAL.Text='' then
   BEGIN
   tz := tzSection ;
   TOBGene.InitValeurs ; TGLibelle.Caption:='' ;
   Case Y_AXE.ItemIndex of
   0 : tz := tzSection ;
   1 : tz := tzSection2 ;
   2 : tz := tzSection3 ;
   3 : tz := tzSection4 ;
   4 : tz := tzSection5 ;
   end ;
   Y_SECTION.ZoomTable:=tz ;
   Exit ;
   END ;
if ((Y_GENERAL.Text=OldGene) and (Not Force)) then Exit ;
if Y_GENERAL.ExisteH<=0 then
   BEGIN
   if Not GChercheCompte(Y_GENERAL,Nil) then TGLibelle.Caption:='' ;
   END ;
if Not TOBGene.SelectDB('"'+Y_GENERAL.Text+'"',Nil,False) then
   BEGIN
   TOBGene.InitValeurs ;
   END else
   BEGIN
   if ((TOBGene.GetValue('G_VENTILABLE')<>'X') and (TOBSection.GetValue('S_SECTION')<>'')) then BEGIN Y_SECTION.Text:='' ; ChargeSection(False) ; END ;
   END ;
TGLibelle.Caption:=TOBGene.GetValue('G_LIBELLE') ;
AfficheLeSoldeUnique ;
END ;

procedure TFConsEcrAna.Y_GENERALExit(Sender: TObject);
begin
ChargeGeneral(False) ;
end;

procedure TFConsEcrAna.ChargeSection ( Force : boolean ) ;
var Q : TQuery ;
BEGIN
if ((TOBSection=Nil) or (csDestroying in ComponentState)) then Exit ;
AfficheLeSoldeUnique ;
  inherited;
if Y_SECTION.Text='' then BEGIN TOBSection.InitValeurs ; TSLibelle.Caption:='' ; Exit ; END ;
if ((Y_SECTION.Text=OldSect) and (Not Force)) then Exit ;
{if Y_SECTION.ExisteH<=0 then
   BEGIN
   if Not GChercheCompte(Y_SECTION,Nil) then TSLibelle.Caption:='' ;
   END ;}
Q := OpenSql('SELECT S_SECTION, S_LIBELLE FROM SECTION WHERE S_SECTION="'+Y_SECTION.text+'"',True ) ;
if Not TOBSection.SelectDB('',Q,False) then
   BEGIN
   TOBSection.InitValeurs ;
   END ;
Ferme(Q) ;
TSLibelle.Caption:=TOBSection.GetValue('S_LIBELLE') ;
AfficheLeSoldeUnique ;
END ;

procedure TFConsEcrAna.Y_SECTIONExit(Sender: TObject);
begin
ChargeSection(False) ;
end;

procedure TFConsEcrAna.Y_GENERALEnter(Sender: TObject);
begin
  inherited;
  OldGene:=Y_GENERAL.Text ;
end;

procedure TFConsEcrAna.Y_SECTIONEnter(Sender: TObject);
begin
  inherited;
OldSect:=Y_SECTION.Text ;
end;

procedure TFConsEcrAna.AfficheLeSoldeUnique ;
Var Gene,Sect,SQL,StC,chD,chC : String ;
    QQ        : TQuery ;
    XD,XC     : Double ;
    sReq      : String ;
BEGIN
Gene:='' ; Sect:='' ; XD:=0 ; XC:=0 ;
if ((TOBSection.GetValue('S_SECTION')<>'') and (Y_SECTION.Text<>'')) then Sect:=TOBSection.GetValue('S_SECTION') ;
if ((TOBGene.GetValue('G_GENERAL')<>'') and (Y_GENERAL.Text<>'')) then Gene:=TOBGene.GetValue('G_GENERAL') ;
if ((Sect<>'') or (Gene<>'')) then
   BEGIN
   if ROppose.ItemIndex=0 then BEGIN chD:='Y_DEBIT' ; chC:='Y_CREDIT' ; END
                          else BEGIN chD:='Y_DEBITEURO' ; chC:='Y_CREDITEURO' ; END ;
   SReq:='SELECT SUM('+chD+'), SUM('+chC+') FROM ANALYTIQ '+RecupWhereCritere(Pages) ;
   QQ:=OpenSQL(sReq,True) ;
   if Not QQ.EOF then
      BEGIN
      XD:=QQ.Fields[0].AsFloat ; XC:=QQ.Fields[1].AsFloat ;
      END ;
   Ferme(QQ) ;
   END ;
AfficheLeSolde(LeSolde,XD,XC) ;
StC:=Copy(Caption,1,Pos(':',Caption)) ;
if ((Y_SECTION.Text<>'') and (Y_GENERAL.Text<>'')) then
   BEGIN
   StC:=StC+' '+Y_GENERAL.Text+'  '+Y_SECTION.Text+'  '+TSLIBELLE.Caption+'   '+LESOLDE.Text ;
   END else
   BEGIN
   if Y_SECTION.Text<>'' then StC:=StC+'  '+Y_SECTION.Text+'  '+TSLIBELLE.Caption+'   '+LESOLDE.Text else
     if Y_GENERAL.Text<>'' then StC:=StC+' '+Y_GENERAL.Text+'  '+TGLIBELLE.Caption+'   '+LESOLDE.Text
                           else StC:=StC ;
   END ;
Caption:=StC ; UpdateCaption(Self) ;
END ;

procedure TFConsEcrAna.GotoEntete ;
Var i : integer ;
BEGIN
if Y_GENERAL.CanFocus then Y_GENERAL.SetFocus else
 if Y_SECTION.CanFocus then Y_SECTION.SetFocus else
    BEGIN
    for i:=0 to Pages.ActivePage.ControlCount-1 do if Pages.ActivePage.Controls[i] is TWinControl then
        if TWinControl(Pages.ActivePage.Controls[i]).CanFocus then
        BEGIN
        TWinControl(Pages.ActivePage.Controls[i]).SetFocus ;
        Break ;
        END ;
    END ;
END ;

procedure TFConsEcrAna.FormCreate(Sender: TObject);
begin
  inherited;
// création des tob pour comptes généraux et sections
TOBGene:=TOB.Create('GENERAUX',Nil,-1) ;
TOBSection:=TOB.Create('SECTION',Nil,-1) ;
Q.Manuel:=True ; ToutAcc:=True ;
MemoStyle:=msBook ; FocusUD:=False ;
end;

procedure TFConsEcrAna.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : boolean ;
begin
  inherited;
Vide:=(Shift=[]) ;
Case Key of
   VK_NEXT  : BEGIN
              if FListe.Focused then Exit ;
              Key:=0 ;
              if Y_GENERAL.Focused then GeneSuivant(True) else if Y_SECTION.Focused then SectSuivant(True) ;
              END ;
   VK_PRIOR : BEGIN
              if FListe.Focused then Exit ;
              Key:=0 ;
              if Y_GENERAL.Focused then GeneSuivant(False) else if Y_SECTION.Focused then SectSuivant(False) ;
              END ;
   VK_F5 : if Vide then
              BEGIN
              Key:=0 ;
              if FListe.Focused then FListeDblClick(Nil) ;
              if Y_GENERAL.Focused then BEGIN if Y_GENERAL.Text='' then Y_GENERAL.DblClick else ChargeGeneral(True) ; END ;
              if Y_SECTION.Focused then BEGIN if Y_SECTION.Text='' then Y_SECTION.DblClick else ChargeSection(True) ; END ;
              END ;
   VK_F6 : if Vide then BEGIN Key:=0 ; if FListe.Focused then GotoEntete else FListe.SetFocus ; END ;
  VK_F11 : BEGIN
           Key:=0 ; PopZ.PopUp(Mouse.CursorPos.X,Mouse.CursorPos.Y) ;
           END ;
   END ;
end;

procedure TFConsEcrAna.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
TOBGene.Free ; TOBGene:=Nil ;
TOBSection.Free ; TOBSection:=Nil ;
PurgePopUp(PopZ) ;
end;

Function TFConsEcrAna.PositionneExo : TExoDate ;
Var LeExo : TExoDate ;
BEGIN
LeExo:=VH^.Entree ;
if Y_EXERCICE.Value=VH^.Encours.Code then LeExo:=VH^.Encours else
 if Y_EXERCICE.Value=VH^.Suivant.Code then LeExo:=VH^.Suivant else
   if Y_EXERCICE.Value=VH^.Precedent.Code then LeExo:=VH^.Precedent ;
Result:=LeExo ;
END ;

{Modif écriture}
procedure TFConsEcrAna.ModifEcrClick(Sender: TObject);
begin
  inherited;
FListeDBLClick(Nil) ;
end;

{Modif compte général}
procedure TFConsEcrAna.ModifGENEClick(Sender: TObject);
Var AA : TActionFiche ;
begin
if TOBGene.GetValue('G_GENERAL')='' then Exit ;
  inherited;
AA:=taModif ; if ((not ExJaiLeDroitConcept(TConcept(ccGenModif),False)) or (Not ToutAcc)) then AA:=taConsult ;
FicheGene(Nil,'',TOBGene.GetValue('G_GENERAL'),AA,0) ;
end;

{Cumuls du compte général}
procedure TFConsEcrAna.CumulsGENEClick(Sender: TObject);
Var LeExo : TExoDate ;
begin
  inherited;
LeExo:=PositionneExo ;
CumulCpteMensuel(fbGene,Y_GENERAL.Text,TOBGene.GetValue('G_LIBELLE'),LeExo) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 26/11/2001
Modifié le ... :   /  /    
Description .. : Remplacement des champs spécifiques aux écritures 
Suite ........ : analytiques par 1=1.
Mots clefs ... :
*****************************************************************}
function SupprimeChampAnalyt ( QModif, Champ : string ) : string ;
var MaChaine : string ;
    PosChamp,PosCote : integer ;
begin
  result := QModif ;
  PosChamp := Pos(Champ, QModif ) ;
  if PosChamp = 0 then exit ;
  while PosChamp <> 0 do
    begin
    // récupération de la condition sue le champ concerné
    MaChaine := Copy(QModif, PosChamp, length(QModif) ) ;
    // suppression de la première "
    MaChaine[Pos('"',MaChaine)] := ' ' ;
    // recherche position de la deuxième "
    PosCote := Pos('"', MaChaine) ;
    // remplacement de la condition sur le champ par 1=1
    system.Delete(QModif,PosChamp,PosCote) ;
    System.Insert('1=1',QModif,PosChamp) ;
    PosChamp := Pos (Champ, QModif ) ;
    end ;
  Result := QModif ;
end ;

{Grand-livre général simple}
procedure TFConsEcrAna.GLGENEClick(Sender: TObject);
Var D1,D2 : TDateTime ;
    Crit  : TCritEDT ;
    QModif : string ;
begin
  inherited;
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=StrToDate(Y_DATECOMPTABLE.text) ; D2:=StrToDate(Y_DATECOMPTABLE_.text) ;
Crit.Exo:=PositionneExo ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neGL ;
InitCritEdt(Crit) ;
Crit.GL.ForceNonCentralisable:=TRUE ;
Crit.Cpt1:=Y_GENERAL.Text ; Crit.Cpt2:=Crit.Cpt1 ;
// récupération des critères sur écritures analytiques et
// adaptation pour écritures comptables
QModif := Q.CRITERES ;
QModif := SupprimeChampAnalyt(QModif,'Y_SECTION') ;
QModif := SupprimeChampAnalyt(QModif,'Y_AXE') ;
QModif := SupprimeChampAnalyt(QModif,'Y_TYPEANALYTIQUE') ;
while Pos('Y_', QModif) > 0 do QModif[Pos('Y_', QModif)]:='E';
Crit.SQLPLUS:=' AND '+QModif+' ' ;
GLGeneralZoom(Crit) ;
end;

procedure TFConsEcrAna.EnabledMenus ;
BEGIN
// Modifs fiches
ModifGene.Enabled:=TOBGene.GetValue('G_GENERAL')<>'' ;
ModifSection.Enabled := Not ((Q.EOF) and (Q.BOF)) ;
// Ecritures
ModifEcr.Enabled:=Not ((Q.EOF) and (Q.BOF)) ;
//Etats
CumulsGene.Enabled := TOBGene.GetValue('G_GENERAL')<>'' ;
GLGene.Enabled := TOBGene.GetValue('G_GENERAL')<>'' ;
CumulSect.Enabled := Not ((Q.EOF) and (Q.BOF)) ;
GLAnalyt.Enabled := Not ((Q.EOF) and (Q.BOF)) ;
END ;

procedure TFConsEcrAna.PopACPTPopup(Sender: TObject);
begin
  inherited;
EnabledMenus ;
end;

procedure TFConsEcrAna.PopAEdtPopup(Sender: TObject);
begin
  inherited;
EnabledMenus ;
end;

procedure TFConsEcrAna.PopZPopup(Sender: TObject);
Var i : integer ;
    TS,TD : TMenuItem ;
    OldCap : String ;
begin
  inherited;
PurgePopUp(PopZ) ; EnabledMenus ; OldCap:='-' ;
for i:=0 to PopACpt.Items.Count-1 do
    BEGIN
    TS:=TMenuItem(PopACpt.Items[i]) ;
    if TS.Enabled then
       BEGIN
       if ((TS.Caption='-') and (OldCap='-')) then Continue ;
       TD:=TMenuItem.Create(PopZ) ; OldCap:=TS.Caption ;
       TD.Caption:=TS.Caption ; TD.OnClick:=TS.OnClick ;
       PopZ.Items.Add(TD) ;
       END ;
    END ;
for i:=0 to PopAEdt.Items.Count-1 do
    BEGIN
    TS:=TMenuItem(PopAEdt.Items[i]) ;
    if TS.Enabled then
       BEGIN
       if ((TS.Caption='-') and (OldCap='-')) then Continue ;
       TD:=TMenuItem.Create(PopZ) ;
       TD.Caption:=TS.Caption ; TD.OnClick:=TS.OnClick ;
       PopZ.Items.Add(TD) ;
       END ;
    END ;
if POPZ.Items.Count>0 then
   BEGIN
   TD:=POPZ.Items[POPZ.Items.Count-1] ;
   if TD.Caption='-' then BEGIN POPZ.Items.Remove(TD) ; TD.Free ; TD:=Nil ; END ;
   END ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 26/11/2001
Modifié le ... :   /  /
Description .. : Recherche du compte général suivant ou précédent
Suite ........ : suivant les critères.
Mots clefs ... :
*****************************************************************}
procedure TFConsEcrAna.GeneSuivant ( Suiv : boolean ) ;
Var Q : TQuery ;
    SQL,CptGene,NewCpt : String ;
BEGIN
CptGene:=Y_GENERAL.Text ; NewCpt:=CptGene ;
if Suiv then SQL:='SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL>"'+CptGene+'"'
        else SQL:='SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL<"'+CptGene+'"' ;
SQL := SQL + ' AND G_VENTILABLE="X" ' ;
Case RDefil.ItemIndex of
   0 : SQL:=SQL+'' ;
   1 : SQL:=SQL+' AND G_TOTALDEBIT-G_TOTALCREDIT<>0 ' ;
   2 : SQL:=SQL+' AND (G_TOTALDEBIT<>0 OR G_TOTALCREDIT<>0) ' ;
   END ;
if Suiv then SQL:=SQL+' ORDER BY G_GENERAL ' else SQL:=SQL+' ORDER BY G_GENERAL DESC ' ;
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then NewCpt:=Q.Fields[0].AsString ;
Ferme(Q) ;
if NewCpt<>CptGene then
   BEGIN
   Y_GENERAL.Text:=NewCpt ; ChargeGeneral(True) ; BChercheClick(Nil) ;
   OldGene:=Y_GENERAL.Text ;
   END ;
END ;

Function TFConsEcrAna.RecupWhereDataType : String ;
Var St : String ;
    ii  : integer ;
    tz  : TZoomTable ;
BEGIN
Result:='' ; St:='' ;
tz:=Y_SECTION.ZoomTable ;
Case tz of
//  tzTToutDebit  : St:='TZTTOUTDEBIT' ;
//  tzTToutCredit : St:='TZTTOUTCREDIT' ;
//  tzTSalarie    : St:='TZTSALARIE' ;
//  tzTiers       : St:='TZTTOUS' ;
  tzSection     : St:='TZSECTION' ;
  else St:='TZTTOUS' ;
  END ;
ii:=TTToNum(St) ; if ii>0 then Result:=V_PGI.DECombos[ii].Where ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 26/11/2001
Modifié le ... :   /  /
Description .. : Recherche de la section suivante ou précédente suivant
Suite ........ : les critères.
Mots clefs ... :
*****************************************************************}
procedure TFConsEcrAna.SectSuivant ( Suiv : boolean ) ;
Var Q : TQuery ;
    SQL,Sect,NewSect,sWhere : String ;
BEGIN
Sect:=Y_SECTION.Text ; NewSect:=Sect ;
if Suiv then SQL:='SELECT S_SECTION FROM SECTION WHERE S_SECTION>"'+Sect+'" '
        else SQL:='SELECT S_SECTION FROM SECTION WHERE S_SECTION<"'+Sect+'" ' ;
sWhere:=RecupWhereDataType ; if sWhere<>'' then SQL:=SQL+' AND '+sWhere ;
Case RDefil.ItemIndex of
   0 : SQL:=SQL+'' ;
   1 : SQL:=SQL+' AND S_TOTALDEBIT-S_TOTALCREDIT<>0 ' ;
   2 : SQL:=SQL+' AND (S_TOTALDEBIT<>0 OR S_TOTALCREDIT<>0) ' ;
   END ;
if Suiv then SQL:=SQL+' ORDER BY S_SECTION ' else SQL:=SQL+' ORDER BY S_SECTION DESC ' ;
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then NewSect:=Q.Fields[0].AsString ;
Ferme(Q) ;
if NewSect<>Sect then
   BEGIN
   Y_SECTION.Text:=NewSect ; ChargeSection(True) ; BChercheClick(Nil) ;
   OldSect:=Y_SECTION.Text ;
   END ;
END ;

procedure TFConsEcrAna.RActionClick(Sender: TObject);
begin
  inherited;
ToutAcc:=(RAction.ItemIndex=1) ;
end;

procedure TFConsEcrAna.ROpposeClick(Sender: TObject);
begin
  inherited;
BChercheClick(Nil) ;
end;

procedure TFConsEcrAna.BNouvRechClick(Sender: TObject);
Var Sect,Gene : String ;
begin
Gene:=Y_GENERAL.Text ; Sect:=Y_SECTION.Text ;
  inherited;
Y_GENERAL.Text:=Gene ; Y_SECTION.Text:=Sect ;
end;

procedure TFConsEcrAna.SectUpClick(Sender: TObject);
begin
  inherited;
SectSuivant(False) ;
end;

procedure TFConsEcrAna.SectDownClick(Sender: TObject);
begin
  inherited;
SectSuivant(True) ;
end;

procedure TFConsEcrAna.GeneUpClick(Sender: TObject);
begin
  inherited;
GeneSuivant(False) ;
end;

procedure TFConsEcrAna.GeneDownClick(Sender: TObject);
begin
  inherited;
GeneSuivant(True) ;
end;

procedure TFConsEcrAna.FormResize(Sender: TObject);
begin
  inherited;
if ((Not IsInside(Self)) and (HMTrad.ActiveResize) and (HMTrad.ResizeDBGrid)) then
   BEGIN
   HMTrad.ResizeDBGridColumns(FListe) ;
   FListe.Refresh ;
   END ;
end;

procedure TFConsEcrAna.CAvecSimuClick(Sender: TObject);
begin
  inherited;
XX_WHEREQUALIF.Text:='Y_QUALIFPIECE="N"' ;
if cAvecSimu.Checked then XX_WHEREQUALIF.Text:='Y_QUALIFPIECE="N" OR Y_QUALIFPIECE="S"' ;
end;

procedure TFConsEcrAna.GLAnalytClick(Sender: TObject);
Var D1,D2 : TDateTime ;
    Crit  : TCritEDT ;
begin
  inherited;
if Y_AXE.Value = '' then
  begin
  HM.Execute(10,Caption,'') ;
  exit ;
  end ;
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=StrToDate(Y_DATECOMPTABLE.text) ; D2:=StrToDate(Y_DATECOMPTABLE_.text) ;
Crit.Gl.Axe:=Y_AXE.Value ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neGL ;
InitCritEdt(Crit) ;
//Crit.GL.ForceNonCentralisable:=TRUE ;
Crit.Cpt1:=Y_SECTION.Text ; Crit.Cpt2:=Crit.Cpt1 ;
Crit.SQLPLUS:=' AND '+Q.CRITERES+' ' ;
GLAnalZoom(Crit) ;
end;

// Cumul de la section
procedure TFConsEcrAna.CumulSectClick(Sender: TObject);
begin
  inherited;
CumulCpteMensuel(AxeToFb(Q.FindField('Y_AXE').asString),Q.FindField('Y_SECTION').AsString,'',VH^.Entree) ;
end;

// création d'une écriture analytique
procedure TFConsEcrAna.SaisieODAClick(Sender: TObject);
Var AA  : TActionFiche ;
    M : RMVT ;
begin
  inherited;
MultiCritereAna(taCreat) ;
BChercheClick(Nil) ;

end;

procedure TFConsEcrAna.Y_AXEChange(Sender: TObject);
begin
  inherited;
  Case Y_AXE.ItemIndex of
  0 : Y_SECTION.ZoomTable := tzSection ;
  1 : Y_SECTION.ZoomTable := tzSection2 ;
  2 : Y_SECTION.ZoomTable := tzSection3 ;
  3 : Y_SECTION.ZoomTable := tzSection4 ;
  4 : Y_SECTION.ZoomTable := tzSection5 ;
  end ;
  Y_SECTION.Text := '' ;
end;

end.

