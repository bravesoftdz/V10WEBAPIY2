unit MulSuivBAP ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB, DBTables, Hqry,
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, Saisie, SaisUtil,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil,
  EncUtil, LettUtil, GenereMP, UTOB, Choix
  ,SaisBor, HRichOLE, TofVerifRib, ADODB, udbxDataset
  ;

Procedure SuivBAP ;

type
  TFMulSuivBAP = class(TFMul)
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
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHEREVIDE: TEdit;
    E_DATEECHEANCE: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    XX_WHEREMP: TEdit;
    BCtrlRib: TToolbarButton97;
    Label14: TLabel;
    E_MODEPAIE: THValComboBox;
    HLabel3: THLabel;
    E_TYPEANOUVEAU: THValComboBox;
    XX_WHERELETTRAGE: TEdit;
    XX_WHERENATCLI: TEdit;
    HLabel1: THLabel;
    E_NOMLOT: THCritMaskEdit;
    HLabel5: THLabel;
    E_NOMLOT_: THCritMaskEdit;
    E_BANQUEPREVI_: THCritMaskEdit;
    HLabel8: THLabel;
    E_BANQUEPREVI: THCritMaskEdit;
    HLabel6: THLabel;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCritMaskEdit;
    HLabel4: THLabel;
    E_AUXILIAIRE_: THCritMaskEdit;
    HLabel7: THLabel;
    E_GENERAL: THCritMaskEdit;
    HLabel9: THLabel;
    E_GENERAL_: THCritMaskEdit;
    HLabel10: THLabel;
    E_SUIVDEC: THValComboBox;
    PCircuit: TTabSheet;
    Bevel7: TBevel;
    g1: TGroupBox;
    HLabel11: THLabel;
    HLabel12: THLabel;
    CircuitSource: THValComboBox;
    LotSource: THCritMaskEdit;
    g2: TGroupBox;
    HLabel13: THLabel;
    HLotDest: THLabel;
    CircuitDest: THValComboBox;
    Label3: TLabel;
    cBAP: TCheckBox;
    cModeleBAP: THValComboBox;
    LotDest: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure BCtrlRibClick(Sender: TObject);
    procedure FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
    procedure E_NOMLOTChange(Sender: TObject);
    procedure BChercheClick(Sender: TObject); override;
    procedure FListeRowEnter(Sender: TObject);
    procedure DessineCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
  private    { Déclarations privées }
    GeneOpe : String ;
    procedure InitCriteres ;
    procedure QueLesLCR ;
    procedure ValideAcc ;
public
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  UtilPGI;

Procedure SuivBAP ;
var X : TFMulSuivBAP ;
    PP : THPanel ;
begin
PP:=FindInsidePanel ;
X:=TFMulSuivBAP.Create(Application) ;
X.Q.Manuel:=True ;
X.FNomFiltre:='CPSUIVBAP' ; X.Q.Liste:='CPSUIVBAP' ;
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
procedure TFMulSuivBAP.FormCreate(Sender: TObject);
begin
  inherited;
MemoStyle:=msBook ;
end;

procedure TFMulSuivBAP.FormShow(Sender: TObject);
begin
InitCriteres ;
QuelesLCR ;
  inherited;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
XX_WHEREVIDE.Text:='' ;
end;

(*============================ INITIALISATIONS ===============================*)
procedure TFMulSuivBAP.InitCriteres ;
//Var i : integer ;
BEGIN
//LibellesTableLibre(PzLibre,'TE_TABLE','E_TABLE','E') ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   E_EXERCICE.Value:=VH^.CPExoRef.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   E_EXERCICE.Value:=VH^.Entree.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.Entree.Deb) ;
   E_DATECOMPTABLE_.Text:=DateToStr(VH^.Entree.Fin) ;
   END ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
//E_DEVISE.Value:=V_PGI.DevisePivot ;
GeneOpe:='' ;
END ;

(*================================ CRITERES ==================================*)
procedure TFMulSuivBAP.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFMulSuivBAP.QueLesLCR ;
begin
  inherited;
XX_WHEREMP.Text:='' ;
end;

procedure TFMulSuivBAP.FListeDblClick(Sender: TObject);
Var sMode : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
sMode:=Q.FindField('E_MODESAISIE').AsString ;
if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q,TypeAction)
                                  else TrouveEtLanceSaisie(Q,taConsult,'N') ;
end;

procedure TFMulSuivBAP.ValideAcc ;
Var SW : String ;
    Nb     : integer ;
BEGIN
SW:='E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
   +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
   +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
   +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
   +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
   +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
   +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString
   +' AND E_QUALIFPIECE="N" ';
Nb:=ExecuteSQL('UPDATE ECRITURE SET E_CODEACCEPT="'+GeneOpe+'" WHERE '+SW) ;
if Nb<>1 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TFMulSuivBAP.BOuvrirClick(Sender: TObject);
Var Ope : String ;
    i   : integer ;
begin
GeneOpe:='' ;
Ope:=Choisir(HM.Mess[0],'COMMUN','CO_LIBELLE','CO_CODE','CO_TYPE="TLC" AND CO_CODE<>"NON" AND CO_CODE<>"BOR"','') ;
if Ope='' then Exit else
   BEGIN
   if HM.Execute(2,caption,'')<>mrYes then Exit ; 
   GeneOpe:=Ope ;
   END ;
if Not FListe.AllSelected then
   BEGIN
   for i:=0 to FListe.NbSelected-1 do
       BEGIN
       FListe.GotoLeBookmark(i) ;
       if Transactions(ValideAcc,1)<>oeOk then
          BEGIN
          MessageAlerte(HM.Mess[1]) ;
          Break ;
          END ;
       END ;
   END else
   BEGIN
   Q.First ;
   While Not Q.EOF do
      BEGIN
      if Transactions(ValideAcc,1)<>oeOk then
         BEGIN
         MessageAlerte(HM.Mess[1]) ;
         Break ;
         END ;
      Q.Next ;
      END ;
   END ;
BChercheClick(Nil) ;
end;

procedure TFMulSuivBAP.BCtrlRibClick(Sender: TObject);
Var
  StWRib : String ;
  i : Integer;
begin
  inherited;
  StWRib := RecupWhereCritere(Pages) ;
  If (StWRib = '') Then Exit;
  if ((Not FListe.AllSelected) and (FListe.NbSelected>0) and (FListe.NbSelected<100)) then begin  // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
    // Rajoute une clause au WHERE
    StWRib := StWRib+' AND (';
    for i:=0 to FListe.NbSelected-1 do begin
      FListe.GotoLeBookmark(i) ;
      StWRib := StWRib +' (E_NUMEROPIECE='+ Q.FindField('E_NUMEROPIECE').AsString +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
    end;
    // Efface le dernier OR et rajoute ')'
    delete(StWRib,length(StWRib)-2,3);
    StWRib := StWRib +')';
  end;
  If StWRib<>'' Then CPLanceFiche_VerifRib('WHERE='+StWRib);
end;

procedure TFMulSuivBAP.FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
Var sRIB : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
if ((Field.FieldName='E_AUXILIAIRE') and (Q.FindField('E_RIB')<>Nil)) then
   BEGIN
   sRIB:=Q.FindField('E_RIB').AsString ; if sRib<>'' then Exit ;
   FListe.Canvas.Brush.Color:= clRed ; FListe.Canvas.Brush.Style:=bsSolid ; FListe.Canvas.Pen.Color:=clRed ;
   FListe.Canvas.Pen.Mode:=pmCopy ; FListe.Canvas.Pen.Width:= 1 ;
   FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
   END ;
end;

procedure TFMulSuivBAP.E_NOMLOTChange(Sender: TObject);
begin
  inherited;
If (E_NOMLOT.Text<>E_NOMLOT_.Text) And (E_NOMLOT.Text<>'') And (E_NOMLOT_.Text<>'') Then LotSource.Text:='Lots multiples' Else
  BEGIN
  If (E_NOMLOT.Text<>'') Then LotSource.Text:=E_NOMLOT.Text Else LotSource.Text:=E_NOMLOT_.Text
  END ;

end;

procedure TFMulSuivBAP.BChercheClick(Sender: TObject);
begin
  inherited;
  FListeRowEnter(nil); // 10930
end;

procedure TFMulSuivBAP.FListeRowEnter(Sender: TObject);
begin
// 10930
  inherited;
  VH^.MPModifFaite:=FALSE ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL Then VH^.MPPop.MPExoPop:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
  If Q.FindField('E_GENERAL')<>NIL Then VH^.MPPop.MPGenPop:=Q.FindField('E_GENERAL').AsString ;
  If Q.FindField('E_AUXILIAIRE')<>NIL Then VH^.MPPop.MPAuxPop:=Q.FindField('E_AUXILIAIRE').AsString ;
  If Q.FindField('E_JOURNAL')<>NIL Then VH^.MPPop.MPJalPop:=Q.FindField('E_JOURNAL').AsString ;
  If Q.FindField('E_NUMEROPIECE')<>NIL Then VH^.MPPop.MPNumPop:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  If Q.FindField('E_NUMLIGNE')<>NIL Then VH^.MPPop.MPNumLPop:=Q.FindField('E_NUMLIGNE').AsInteger ;
  If Q.FindField('E_NUMECHE')<>NIL Then VH^.MPPop.MPNumEPop:=Q.FindField('E_NUMECHE').AsInteger ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL Then VH^.MPPop.MPDatePop:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 26/10/2005
Modifié le ... :   /  /    
Description .. : Surcharge de la méthode de Mul.pas pour remplacer le 
Suite ........ : OnDrawDataCell qui ne fonctionne pas
Mots clefs ... : 
*****************************************************************}
procedure TFMulSuivBAP.DessineCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
Var sRIB : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
if ((Column.Field.FieldName='E_AUXILIAIRE') and (Q.FindField('E_RIB')<>Nil)) then
   BEGIN
   sRIB:=Q.FindField('E_RIB').AsString ; if sRib<>'' then Exit ;
   FListe.Canvas.Brush.Color:= clRed ; FListe.Canvas.Brush.Style:=bsSolid ; FListe.Canvas.Pen.Color:=clRed ;
   FListe.Canvas.Pen.Mode:=pmCopy ; FListe.Canvas.Pen.Width:= 1 ;
   FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
   END ;
end;

end.

