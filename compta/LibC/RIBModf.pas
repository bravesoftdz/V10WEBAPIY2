unit RIBModf;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, Hctrls, Mask, ComCtrls, DB,
  DBTables, Hqry, ParamDBG, PrintDBG, Ent1, HEnt1, Paramdat,
  Saisutil, Saiscomm, Saisie, hmsgbox, Filtre, HDB, Menus,
  Hcompte, HSysMenu, HTB97, HPanel, UiUtil,EcheMPA, ADODB, udbxDataset ;

type
  TFRIBModf = class(TForm)
    Pages: TPageControl;
    Princ: TTabSheet;
    Bevel1: TBevel;
    HLabel4: THLabel;
    HLabel1: THLabel;
    TE_EXERCICE: THLabel;
    HLabel3: THLabel;
    HLabel6: THLabel;
    E_GENERAL: THCpteEdit;
    E_AUXILIAIRE: THCpteEdit;
    E_EXERCICE: THValComboBox;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    Complements: TTabSheet;
    Label12: THLabel;
    Bevel2: TBevel;
    E_REFINTERNE: THCritMaskEdit;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    BChercher: TToolbarButton97;
    HPB: TToolWindow97;
    BReduire: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BRecherche: TToolbarButton97;
    BParamListe: TToolbarButton97;
    GL: THDBGrid;
    HLabel8: THLabel;
    E_DEVISE: THValComboBox;
    HLabel10: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    HLabel7: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    FindMvt: TFindDialog;
    QEche: THQuery;
    SEche: TDataSource;
    E_ECHE: TEdit;
    E_NUMECHE: THCritMaskEdit;
    XX_WHERE: TEdit;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    BZoomPiece: TToolbarButton97;
    HME: THMsgBox;
    E_TRESOLETTRE: THCritMaskEdit;
    HMTrad: THSystemMenu;
    TE_QUALIFPIECE: THLabel;
    E_QUALIFPIECE: THValComboBox;
    Label1: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    Label2: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    Dock: TDock97;
    Dock971: TDock97;
    MP_CODEACCEPT: THMultiValComboBox;
    TMP_CODEACCEPT: THLabel;
    PModifS: TTabSheet;
    Bevel3: TBevel;
    MPNew: THValComboBox;
    DateEcheNew: THCritMaskEdit;
    RIBNEW: TEdit;
    ZoomRib: TToolbarButton97;
    Label3: TLabel;
    Label4: TLabel;
    BModifSerie: TToolbarButton97;
    CMSMP: TCheckBox;
    CMSDateEche: TCheckBox;
    CMSRIB: TCheckBox;
    bSelectAll: TToolbarButton97;
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure FindMvtFind(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_DATEECHEANCEKeyPress(Sender: TObject; var Key: Char);
    procedure BZoomPieceClick(Sender: TObject);
    procedure GLDblClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GLKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure RIBNEWKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ZoomRibClick(Sender: TObject);
    procedure BModifSerieClick(Sender: TObject);
    procedure E_AUXILIAIREChange(Sender: TObject);
    procedure CMSMPClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
  private
    WMinX,WMinY : Integer ;
    LastQualif  : String ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    procedure InitCriteres ;
    procedure ClickModifMPA ;
    procedure ClickModifRIB ;
    procedure ZoomSurRib ;
    procedure ClickPourMS ;
    procedure ModifSerieChamp ;
  public
    FindFirst : boolean ;
  end;

procedure ModifLesRib ;

implementation

{$R *.DFM}

Uses UtilPgi ;

procedure ModifLesRib ;
Var X  : TFRIBModf ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture','nrBatch','nrLettrage','nrEnca','nrDeca'],True,'nrSaisieModif') then Exit ;
X:=TFRIBModf.Create(Application) ;
X.QEche.Manuel:=True ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
     _Bloqueur('nrSaisieModif',False) ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure TFRIBModf.BCreerFiltreClick(Sender: TObject);
begin NewFiltre('MODIFRIB',FFiltres,Pages) ; end;

procedure TFRIBModf.BSaveFiltreClick(Sender: TObject);
begin SaveFiltre('MODIFRIB',FFiltres,Pages) ; end;

procedure TFRIBModf.BDelFiltreClick(Sender: TObject);
begin DeleteFiltre('MODIFRIB',FFiltres) ; end;

procedure TFRIBModf.BRenFiltreClick(Sender: TObject);
begin RenameFiltre('MODIFRIB',FFiltres) ; end;

procedure TFRIBModf.BNouvRechClick(Sender: TObject);
begin VideFiltre(FFiltres,Pages) ; InitCriteres ; end;

procedure TFRIBModf.FFiltresChange(Sender: TObject);
begin LoadFiltre('MODIFRIB',FFiltres,Pages) ; end;

procedure TFRIBModf.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFRIBModf.BImprimerClick(Sender: TObject);
begin PrintDBGrid (GL,Pages,Caption,''); end;

procedure TFRIBModf.BChercherClick(Sender: TObject);
begin
LastQualif:=E_QUALIFPIECE.Value ;
if ((E_GENERAL.Text='') and (E_AUXILIAIRE.Text='') and (Sender<>Nil)) then if HME.Execute(1,'','')<>mrYes then Exit ;
if ((VH^.ExoV8.Code<>'') and (StrToDate(E_DATECOMPTABLE.Text)<VH^.ExoV8.Deb)) then E_DATECOMPTABLE.Text:=DateToStr(VH^.ExoV8.Deb) ;
QEche.UpdateCriteres ;
CentreDBGrid(GL) ;
end;

procedure TFRIBModf.BParamListeClick(Sender: TObject);
begin ParamListe(QEche.Liste,GL,QEche) ; end;

procedure TFRIBModf.BRechercheClick(Sender: TObject);
begin FindFirst:=True ; FindMvt.Execute ; end;

procedure TFRIBModf.FindMvtFind(Sender: TObject);
begin Rechercher(GL,FindMvt,FindFirst) ; end;

procedure TFRIBModf.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFRIBModf.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFRIBModf.E_EXERCICEChange(Sender: TObject);
begin ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ; end;

procedure TFRIBModf.E_DATEECHEANCEKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFRIBModf.FormShow(Sender: TObject);
begin
DateEcheNew.Text:=DateToStr(V_PGI.DateEntree) ;
E_QUALIFPIECE.Value:='N' ; LastQualif:='N' ;
E_EXERCICE.Value:=VH^.Entree.Code ; E_EXERCICEChange(Nil) ;
E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ; E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
ChargeFiltre('MODIFRIB',FFiltres,Pages) ;
if ((E_JOURNAL.Value='') and (E_JOURNAL.Values.Count>0)) then
   BEGIN
   if Not E_JOURNAL.Vide then E_JOURNAL.Value:=E_JOURNAL.Values[0] else
    if E_JOURNAL.Values.Count>1 then E_JOURNAL.Value:=E_JOURNAL.Values[1] ;
   END ;
QEche.Liste:='MODIFRIBS' ; 
QEche.Manuel:=FALSE ; CentreDBGrid(GL) ;
BChercherClick(Nil) ;
end;

procedure TFRIBModf.BZoomPieceClick(Sender: TObject);
begin
TrouveEtLanceSaisie(QEche,taConsult,LastQualif) ;
end;


procedure TFRIBModf.GLDblClick(Sender: TObject);
begin
ClickModifMPA ;
end;

procedure TFRIBModf.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then E_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else E_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
END ;

procedure TFRIBModf.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFRIBModf.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFRIBModf.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFRIBModf.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;

end;

procedure TFRIBModf.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   BEGIN
   _Bloqueur('nrSaisieModif',False) ;
   Action:=caFree ;
   END ;
end;

procedure TFRIBModf.ClickModifRIB ;
Var M : RMVT ;
    Q : TQuery ;
    Trouv : boolean ;
    TAN   : String3 ;
    RIB,Aux,OldRib : String ;
    k     : integer ;
    Coll  : String ;
    RJal,RExo,RQualif : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
begin
if QEche.EOF then Exit ;
RJal:=QEche.FindField('E_JOURNAL').AsString ;
RExo:=QuelExo(DateToStr(QEche.FindField('E_DATECOMPTABLE').AsDateTime)) ;
RQualif:=LastQualif ;
RDate:=QEche.FindField('E_DATECOMPTABLE').AsDateTime ;
RNumP:=QEche.FindField('E_NUMEROPIECE').AsInteger ;
RNumL:=QEche.FindField('E_NUMLIGNE').AsInteger ;
RNumEche:=QEche.FindField('E_NUMECHE').AsInteger ;
Q:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+QEche.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(QEche.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(QEche.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_QUALIFPIECE="'+LastQualif+'"'
          +' AND E_NUMEROPIECE='+QEche.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+QEche.FindField('E_NUMLIGNE').AsString
          +' AND E_NUMECHE='+QEche.FindField('E_NUMECHE').AsString,FALSE) ;
Trouv:=Not Q.EOF ;
if Trouv then
   BEGIN
   M:=MvtToIdent(Q,fbGene,True) ;
   RIB:=Q.FindField('E_RIB').AsString ;
   TAN:=Q.FindField('E_ECRANOUVEAU').AsString ;
   Aux:=Q.FindField('E_AUXILIAIRE').AsString ;
   If Aux='' Then Trouv:=FALSE ;
   if TAN='OAN' then
      BEGIN
      if M.CodeD<>V_PGI.DevisePivot then BEGIN HME.Execute(0,'','') ; Trouv:=FALSE ; END ;
      if ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then BEGIN HME.Execute(0,'','') ; Trouv:=FALSE ; END ;
      END ;
   END ;
If Trouv Then
   BEGIN
   OldRIB:=RIB ; ModifLeRIB(Rib,Aux) ;
   If RIB<>OldRib Then
      BEGIN
      Q.Edit ;
      Q.FindField('E_RIB').AsString:=RIB ;
      Q.Post ;
      END ;
   Ferme(Q) ;
   Application.ProcessMessages ; BChercherClick(Nil) ;
   QEche.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                VarArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[])
   END Else Ferme(Q) ;
end;


procedure TFRIBModf.ClickModifMPA ;
Var M : RMVT ;
    Q : TQuery ;
    Trouv,OkModif : boolean ;
    TAN   : String3 ;
    RIB,Aux,OldRib : String ;
    k     : integer ;
    Coll  : String ;
    RJal,RExo,RQualif : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
    EcheMPA,OldEcheMPA : T_EcheMPA ;
begin
if QEche.EOF then Exit ;
RJal:=QEche.FindField('E_JOURNAL').AsString ;
RExo:=QuelExo(DateToStr(QEche.FindField('E_DATECOMPTABLE').AsDateTime)) ;
RQualif:=LastQualif ;
RDate:=QEche.FindField('E_DATECOMPTABLE').AsDateTime ;
RNumP:=QEche.FindField('E_NUMEROPIECE').AsInteger ;
RNumL:=QEche.FindField('E_NUMLIGNE').AsInteger ;
RNumEche:=QEche.FindField('E_NUMECHE').AsInteger ;
Q:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+QEche.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(QEche.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(QEche.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_QUALIFPIECE="'+LastQualif+'"'
          +' AND E_NUMEROPIECE='+QEche.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+QEche.FindField('E_NUMLIGNE').AsString
          +' AND E_NUMECHE='+QEche.FindField('E_NUMECHE').AsString,FALSE) ;
Trouv:=Not Q.EOF ;
if Trouv then
   BEGIN
   M:=MvtToIdent(Q,fbGene,True) ;
   RIB:=Q.FindField('E_RIB').AsString ;
   TAN:=Q.FindField('E_ECRANOUVEAU').AsString ;
   Aux:=Q.FindField('E_AUXILIAIRE').AsString ;
   If Aux='' Then Trouv:=FALSE ;
   if TAN='OAN' then
      BEGIN
      if M.CodeD<>V_PGI.DevisePivot then BEGIN HME.Execute(0,'','') ; Trouv:=FALSE ; END ;
      if ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then BEGIN HME.Execute(0,'','') ; Trouv:=FALSE ; END ;
      END ;
   END ;
If Trouv Then
   BEGIN
   EcheMPA.DateEche:=Q.Findfield('E_DATEECHEANCE').AsDateTime ;
   EcheMPA.DateComptable:=Q.Findfield('E_DATECOMPTABLE').AsDateTime ;
   EcheMPA.ModePaie:=Q.Findfield('E_MODEPAIE').AsString ;
   EcheMPA.Jal:=Q.Findfield('E_JOURNAL').AsString ;
   EcheMPA.NatP:=Q.Findfield('E_NATUREPIECE').AsString ;
   EcheMPA.Aux:=Aux ;
   EcheMPA.Rib:=RIB ;
   EcheMPA.Montant:=Q.Findfield('E_DEBIT').AsFloat+Q.Findfield('E_CREDIT').AsFloat ;
   EcheMPA.NumP:=Q.Findfield('E_NUMEROPIECE').AsInteger ;
   EcheMPA.RefExterne:=Q.Findfield('E_REFEXTERNE').AsString ;
   OldEcheMPA:=EcheMPA ;
   OkModif:=FALSE ;
   If ModifUneEcheanceMPA(EcheMPA) Then
      BEGIN
      Q.Edit ;
      If EcheMPA.ModePaie<>OldEcheMPA.ModePaie Then BEGIN Q.FindField('E_MODEPAIE').AsString:=EcheMPA.MODEPAIE ; OkModif:=TRUE ; END ;
      If EcheMPA.DateEche<>OldEcheMPA.DateEche Then BEGIN Q.FindField('E_DATEECHEANCE').AsDateTime:=EcheMPA.DateEche ; OkModif:=TRUE ; END ;
      If EcheMPA.Rib<>OldEcheMPA.Rib Then BEGIN Q.FindField('E_RIB').AsString:=EcheMPA.Rib ; OkModif:=TRUE ; END ;
      If EcheMPA.RefExterne<>OldEcheMPA.RefExterne Then BEGIN Q.FindField('E_REFEXTERNE').AsString:=EcheMPA.RefExterne ; OkModif:=TRUE ; END ;
      // GG COM
      If OkSynChro Then Q.FindField('E_PAQUETREVISION').AsInteger:=1 ;
      Q.Post ;
      END ;
   Ferme(Q) ;
   Application.ProcessMessages ; If OkModif Then BChercherClick(Nil) ;
   QEche.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                VarArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[])
   END Else Ferme(Q) ;
end;

procedure TFRIBModf.GLKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var Vide : Boolean ;
begin
Vide:=(Shift=[]) ;
if Vide AND (Key=VK_F5) then BEGIN Key:=0 ; ClickModifMPA ; END ;
if (ssCtrl in Shift) AND (Key=VK_F5) then BEGIN Key:=0 ; ClickModifRIB ; END ;
end;

procedure TFRIBModf.ZoomSurRib ;
Var Rib,Aux : String ;
begin
Rib:='' ; Aux:=E_AUXILIAIRE.Text ;
If ModifLeRIB(Rib,Aux) Then RibNew.Text:=Rib ;
end;

procedure TFRIBModf.RIBNEWKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : Boolean ;
begin
Vide:=(Shift=[]) ;
if Vide AND (Key=VK_F5) then BEGIN Key:=0 ; ZoomSurRIB ; END ;
Key:=0 ;
end;

procedure TFRIBModf.ZoomRibClick(Sender: TObject);
begin
ZoomSurRib ;
end;

procedure TFRIBModf.ClickPourMS ;
Var OkMS : Boolean ;
BEGIN
OkMS:=CMSMP.Checked or CMSDateEche.Checked Or CMSRib.Checked ;
BValider.Visible:=Not OkMS ; BModifSerie.Visible:=OkMS ;
MPNew.Enabled:=CMSMP.Checked ; DateEcheNew.Enabled:=CMSDateEche.Checked ;
RIBNew.Enabled:=CMSRib.Checked ;
If OkMS Then
   BEGIN
   END Else
   BEGIN
   END ;
If Not CMSRib.Checked Then RIBNew.Text:='' ;
If Not CMSMP.Checked Then MPNew.ItemIndex:=-1 ;
GL.MultiSelection:=OkMS ; BSelectAll.Visible:=OkMS ;
END ;

procedure TFRIBModf.ModifSerieChamp ;
Var RJal,RExo,RQualif : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
    What,SQL : String ;
BEGIN
RJal:=QEche.FindField('E_JOURNAL').AsString ;
RExo:=QuelExo(DateToStr(QEche.FindField('E_DATECOMPTABLE').AsDateTime)) ;
RQualif:=LastQualif ;
RDate:=QEche.FindField('E_DATECOMPTABLE').AsDateTime ;
RNumP:=QEche.FindField('E_NUMEROPIECE').AsInteger ;
RNumL:=QEche.FindField('E_NUMLIGNE').AsInteger ;
RNumEche:=QEche.FindField('E_NUMECHE').AsInteger ;
What:='' ;
If CMSMP.Checked Then What:=What+',E_MODEPAIE="'+MPNew.Value+'" ' ;
If CMSDateEche.Checked Then What:=What+',E_DATEECHEANCE="'+USDATE(DateEcheNew)+'" ' ;
If CMSRIB.Checked Then What:=What+',E_RIB="'+RibNew.Text+'" ' ;
If Trim(What)='' Then Exit ;
// GG COM
If OkSynChro Then What:=What+',E_PAQUETREVISION=1 ' ;
What:=Copy(What,2,Length(What)-1) ;
SQL:='UPDATE Ecriture SET '+What
          +'where E_JOURNAL="'+QEche.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(QEche.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(QEche.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_QUALIFPIECE="'+LastQualif+'"'
          +' AND E_NUMEROPIECE='+QEche.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+QEche.FindField('E_NUMLIGNE').AsString
          +' AND E_NUMECHE='+QEche.FindField('E_NUMECHE').AsString ;
ExecuteSQL(SQL) ;
END ;

procedure TFRIBModf.BModifSerieClick(Sender: TObject);
Var BM : TBookmark;
    i : Integer ;
    What,SQL : String ;
begin
If CMSMP.Checked And (MPNEW.Text='') Then BEGIN HME.Execute(2,'','') ; Pages.ActivePage:=PModifS ; MPNew.SetFocus ; Exit ; END ;
If CMSDateEche.Checked And (Not IsValidDate(DateEcheNew.Text) ) Then BEGIN HME.Execute(3,'','') ; Pages.ActivePage:=PModifS ; DateEcheNew.SetFocus ; Exit ;END ;
If CMSRib.Checked And (RibNew.Text='') Then BEGIN HME.Execute(3,'','') ; Pages.ActivePage:=PModifS ; RIBNew.SetFocus ; Exit ; END ;
if (GL.NbSelected>0) or (GL.AllSelected) then
   BEGIN
   BM:=NIL ; BM:=QEche.GetBookmark;
   if GL.AllSelected then
     BEGIN
     QEche.DisableControls ;
     QEche.First ;
     While not QEche.Eof do
       BEGIN
       If Transactions(ModifSerieChamp,2)<>oeOk Then ;
       QEche.Next ;
       END ;
     QEche.EnableControls ;
     END else
     BEGIN
     for i:=0 to GL.NbSelected-1 do
      BEGIN
      GL.GotoLeBookMark(i) ;
      If Transactions(ModifSerieChamp,2)<>oeOk Then ;
      END ;
     END ;
   QEche.GotoBookmark(BM); QEche.FreeBookmark(BM);
   BChercherClick(NIL) ;
   END ;
end;

procedure TFRIBModf.E_AUXILIAIREChange(Sender: TObject);
begin
CMSRIB.Enabled:=Trim(E_AUXILIAIRE.Text)<>'' ; ClickPourMS ;
end;

procedure TFRIBModf.CMSMPClick(Sender: TObject);
begin
ClickPourMS ;
end;

procedure TFRIBModf.bSelectAllClick(Sender: TObject);
begin
GL.AllSelected:=Not GL.AllSelected ;
end;

end.
