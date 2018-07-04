unit EcheModf;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, Hctrls, Mask, ComCtrls, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, ParamDBG, PrintDBG, Ent1, HEnt1, Paramdat,
  Saisutil, Saiscomm, Saisie, EcheUnit, hmsgbox, UObjFiltres {YMO 15/05/2006 Nouvelle gestion des filtres}, HDB, Menus,
  Hcompte, HSysMenu, HTB97, HPanel, UiUtil, ADODB ;

procedure ModifEcheances ( smp : TSuiviMP = smpAucun ) ;

type
  TFEcheModf = class(TForm)
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
    HLabel8: THLabel;
    E_DEVISE: THValComboBox;
    HLabel10: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    HLabel7: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    FindMvt: THFindDialog;
    QEche: THQuery;
    SEche: TDataSource;
    E_ETATLETTRAGE: TEdit;
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
    E_ANA: THCritMaskEdit;
    XX_WHERENATTIERS: TEdit;
    XX_WHEREVIDE: TEdit;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    GL: THDBGrid;
   
    
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
    procedure FormCreate(Sender: TObject);   
    procedure BAideClick(Sender: TObject);  
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure PositionneWhere ;
  private
    ObjFiltre : TObjFiltre;  
    WMinX,WMinY : Integer ;
    LastQualif  : String ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
//    procedure InitCriteres ;
  public
    FindFirst : boolean ;
    smp       : TSuiviMP ;
  end;


implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  UtilPgi ;

procedure ModifEcheances ( smp : TSuiviMP = smpAucun ) ;
Var X  : TFEcheModf ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture','nrBatch','nrLettrage','nrEnca','nrDeca'],True,'nrSaisieModif') then Exit ;
PP:=FindInsidePanel ;
X:=TFEcheModf.Create(Application) ;
X.QEche.Manuel:=True ; X.smp:=smp ;
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

procedure TFEcheModf.BFermeClick(Sender: TObject);
begin
  //SG6 04.03.05 FQ 15443
  Close ;
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
end;

procedure TFEcheModf.BImprimerClick(Sender: TObject);
begin PrintDBGrid (GL,Pages,Caption,''); end;

procedure TFEcheModf.BChercherClick(Sender: TObject);
begin
LastQualif:=E_QUALIFPIECE.Value ;
if ((E_GENERAL.Text='') and (E_AUXILIAIRE.Text='') and (Sender<>Nil)) then if HME.Execute(1,'','')<>mrYes then Exit ;
if ((VH^.ExoV8.Code<>'') and (StrToDate(E_DATECOMPTABLE.Text)<VH^.ExoV8.Deb)) then E_DATECOMPTABLE.Text:=DateToStr(VH^.ExoV8.Deb) ;
QEche.UpdateCriteres ;
HMTrad.ResizeDBGridColumns(GL) ;
CentreDBGrid(GL) ;
end;

procedure TFEcheModf.BParamListeClick(Sender: TObject);
begin ParamListe(QEche.Liste,GL,QEche) ; end;

procedure TFEcheModf.BRechercheClick(Sender: TObject);
begin FindFirst:=True ; FindMvt.Execute ; end;

procedure TFEcheModf.FindMvtFind(Sender: TObject);
begin Rechercher(GL,FindMvt,FindFirst) ; end;

procedure TFEcheModf.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFEcheModf.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFEcheModf.E_EXERCICEChange(Sender: TObject);
begin ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ; end;

procedure TFEcheModf.E_DATEECHEANCEKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFEcheModf.PositionneWhere ;
BEGIN
if smp=smpAucun then Exit ;
if smp=smpEncTous then
   BEGIN
   XX_WHERENATTIERS.Text:='T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD"' ;
   E_AUXILIAIRE.ZoomTable:=tztToutDebit ;
   E_GENERAL.ZoomTable:=tzgCollClient ;
   END else
   BEGIN
   XX_WHERENATTIERS.Text:='T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC"' ;
   E_AUXILIAIRE.ZoomTable:=tztToutCredit ;
   E_GENERAL.ZoomTable:=tzgCollFourn ;
   END ;
END ;

procedure TFEcheModf.FormShow(Sender: TObject);
begin
E_QUALIFPIECE.Value:='N' ; LastQualif:='N' ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   E_EXERCICE.Value:=VH^.CPExoRef.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   E_EXERCICE.Value:=VH^.Entree.Code ; E_EXERCICEChange(Nil) ;
   E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
   E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
PositionneEtabUser(E_ETABLISSEMENT) ;

if ((E_JOURNAL.Value='') and (E_JOURNAL.Values.Count>0)) then
   if smp=smpAucun then 
   BEGIN
   if Not E_JOURNAL.Vide then E_JOURNAL.Value:=E_JOURNAL.Values[0] else
    if E_JOURNAL.Values.Count>1 then E_JOURNAL.Value:=E_JOURNAL.Values[1] ;
   END ;
if smp=smpAucun then QEche.Liste:='MODIFECHE' else QEche.Liste:='CPMODIFECHEMP' ;
PositionneWhere ;
QEche.Manuel:=FALSE ; CentreDBGrid(GL) ;
BChercherClick(Nil) ;
XX_WHEREVIDE.Text:='' ;
//YMO 15/05/2006 Nouvelle gestion des filtres
ObjFiltre.Charger;
end;

procedure TFEcheModf.BZoomPieceClick(Sender: TObject);
begin
TrouveEtLanceSaisie(QEche,taConsult,LastQualif) ;
end;

procedure TFEcheModf.GLDblClick(Sender: TObject);
Var M : RMVT ;
    Q : TQuery ;
    Trouv : boolean ;
    EU    : T_ECHEUNIT ;
    TAN   : String3 ;
    k     : integer ;
    Coll  : String ;
begin
if QEche.EOF then Exit ;
Q:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+QEche.FindField('E_JOURNAL').AsString+'"'
          +' AND E_EXERCICE="'+QuelExo(DateToStr(QEche.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
          +' AND E_DATECOMPTABLE="'+USDATETIME(QEche.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
          +' AND E_QUALIFPIECE="'+LastQualif+'"'
          +' AND E_NUMEROPIECE='+QEche.FindField('E_NUMEROPIECE').AsString
          +' AND E_NUMLIGNE='+QEche.FindField('E_NUMLIGNE').AsString
          +' AND E_NUMECHE='+QEche.FindField('E_NUMECHE').AsString,True) ;
Trouv:=Not Q.EOF ;
if Trouv then
   BEGIN
   M:=MvtToIdent(Q,fbGene,True) ; FillChar(EU,Sizeof(EU),#0) ;
   EU.DateEche:=Q.FindField('E_DATEECHEANCE').AsDateTime ; EU.ModePaie:=Q.FindField('E_MODEPAIE').AsString ;
   EU.DebitDEV:=Q.FindField('E_DEBITDEV').AsFloat ; EU.CreditDEV:=Q.FindField('E_CREDITDEV').AsFloat ;
   EU.Debit:=Q.FindField('E_DEBIT').AsFloat ; EU.Credit:=Q.FindField('E_CREDIT').AsFloat ;
   //EU.DebitEuro:=Q.FindField('E_DEBITEURO').AsFloat ;
   //EU.CreditEuro:=Q.FindField('E_CREDITEURO').AsFloat ;
   EU.DEVISE:=Q.FindField('E_DEVISE').AsString ; EU.TauxDEV:=Q.FindField('E_TAUXDEV').AsFloat ;
   EU.DateComptable:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
   EU.DateModif:=Q.FindField('E_DATEMODIF').AsDateTime ;
   EU.ModeSaisie:=Q.FindField('E_MODESAISIE').AsString ;
   //if EuroOK then EU.SaisieEuro:=(Q.FindField('E_SAISIEEURO').AsString='X')
   //          else EU.SaisieEuro:=False ;
   TAN:=Q.FindField('E_ECRANOUVEAU').AsString ;
   {#TVAENC}
   if VH^.OuiTvaEnc then
      BEGIN
      Coll:=Q.FindField('E_GENERAL').AsString ;
      if EstCollFact(Coll) then
         BEGIN
         for k:=1 to 4 do EU.TabTva[k]:=Q.FindField('E_ECHEENC'+IntToStr(k)).AsFloat ;
         EU.TabTva[5]:=Q.FindField('E_ECHEDEBIT').AsFloat ;
         END ;
      END ;
   END ;
Ferme(Q) ;
if Not Trouv then Exit ;
if TAN='OAN' then
   BEGIN
   if M.CodeD<>V_PGI.DevisePivot then BEGIN HME.Execute(0,'','') ; Exit ; END ;
   if ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then BEGIN HME.Execute(0,'','') ; Exit ; END ;
   END ;
if ModifUneEcheance(M,EU) then BEGIN Application.ProcessMessages ; BChercherClick(Nil) ; END ;
end;

(*
procedure TFEcheModf.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then E_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else E_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
END ;
*)

procedure TFEcheModf.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  //YMO 15/05/2006 Nouvelle gestion des filtres
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'MODIFECHE');
  
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFEcheModf.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFEcheModf.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;
    
procedure TFEcheModf.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   BEGIN
   _Bloqueur('nrSaisieModif',False) ;
   END ;
ObjFiltre.Free;   
end;

end.
