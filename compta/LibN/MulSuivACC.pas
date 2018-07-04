unit MulSuivACC ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry,
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, Saisie, SaisUtil,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil,
  EncUtil, LettUtil, GenereMP, UTOB, UtilPGI, Choix
  ,SaisBor, HRichOLE, TofVerifRib,
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  Lookup, ADODB;  //fb 19/06/2006 FQ12487

Procedure SuivAcceptation(Qui : tProfilTraitement) ;

type
  TFMulSuivACC = class(TFMul)
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
    HLabel4: THLabel;
    E_GENERAL: THCritMaskEdit;
    HLabel17: THLabel;
    E_DEVISE: THValComboBox;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCritMaskEdit;
    XX_WHEREMP: TEdit;
    BCtrlRib: TToolbarButton97;
    HLabel1: THLabel;
    E_AUXILIAIRE_: THCritMaskEdit;
    Label14: TLabel;
    E_MODEPAIE: THValComboBox;
    HLabel3: THLabel;
    E_CODEACCEPT: THValComboBox;
    TFLettrage: THLabel;
    FLettrage: THValComboBox;
    XX_WHERELETTRAGE: TEdit;
    XX_WHERENATCLI: TEdit;
    lbl_NumTraite: TLabel;
    E_NUMTRAITECHQ: THCritMaskEdit;
    Label2: TLabel;
    E_NUMTRAITECHQ_: THCritMaskEdit;
    XX_WHEREPROFIL: TEdit;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure BCtrlRibClick(Sender: TObject);
    procedure FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
    procedure FLettrageChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeRowEnter(Sender: TObject);
    procedure BChercheClick(Sender: TObject); override;
    procedure DessineCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
{b fb 19/06/2006 FQ12487}
    procedure E_GENERALElipsisClick(Sender: TObject);
    procedure E_AUXILIAIREElipsisClick(Sender: TObject);
{e fb 19/06/2006 FQ12487}

  private    { Déclarations privées }
    GeneOpe : String ;
    Qui : tProfilTraitement ;
    procedure InitCriteres ;
    procedure QueLesLCR ;
    procedure ValideAcc ;
    procedure ClickModifRib ;

public
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  MulSMPUtil ;

Procedure SuivAcceptation(Qui : tProfilTraitement) ;
var X : TFMulSuivACC ;
    PP : THPanel ;
begin
PP:=FindInsidePanel ;
X:=TFMulSuivACC.Create(Application) ;
X.Q.Manuel:=True ;
X.FNomFiltre:='CPSUIVACC' ; X.Q.Liste:='CPSUIVACC' ;
if PP=Nil then
   BEGIN
    try
     X.Qui:=Qui ;
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
procedure TFMulSuivACC.FormCreate(Sender: TObject);
begin
  inherited;
MemoStyle:=msBook ;   
end;

procedure TFMulSuivACC.FormShow(Sender: TObject);
begin
HelpContext:=3710500 ;
InitCriteres ;
QuelesLCR ;
FLettrage.Value:='NL' ; FLettrage.Visible:=FALSE ; TFLettrage.Visible:=FALSE ; 
  inherited;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
XX_WHEREVIDE.Text:='' ;
end;

(*============================ INITIALISATIONS ===============================*)
procedure TFMulSuivACC.InitCriteres ;
Var //i : integer ;
    St : String ;
BEGIN
LibellesTableLibre(PzLibre,'TE_TABLE','E_TABLE','E') ;
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
E_DEVISE.Value:=V_PGI.DevisePivot ;
GeneOpe:='' ;
XX_WHEREPROFIL.Text:='' ;
St:=WhereProfilUser(Q,Qui) ; XX_WHEREPROFIL.Text:=St ;
END ;

(*================================ CRITERES ==================================*)
procedure TFMulSuivACC.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFMulSuivACC.QueLesLCR ;
Var St,sMode,sCat,sAcc : String ;
    i : integer ;
//    HCAT : THValComboBox ;
begin
  inherited;
XX_WHEREMP.Text:='' ;
for i:=0 to VH^.MPACC.Count-1 do
    BEGIN
    St:=VH^.MPACC[i] ;
    sMode:=ReadtokenSt(St) ; sAcc:=ReadtokenSt(St) ; sCat:=ReadtokenSt(St) ;
    if sCat<>'LCR' then Continue ;
    if ((sAcc<>'BOR') and (sAcc<>'NON')) then XX_WHEREMP.Text:=XX_WHEREMP.Text+' OR E_MODEPAIE="'+sMode+'"' ;
    END ;
if XX_WHEREMP.Text='' then
   BEGIN
   XX_WHEREMP.Text:='E_MODEPAIE="aaa"' ;
   END else // ne rien trouver
   BEGIN
   St:=XX_WHEREMP.Text ; Delete(St,1,4) ; XX_WHEREMP.Text:=St ;
   END ;
end;

procedure TFMulSuivACC.FListeDblClick(Sender: TObject);
Var sMode : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
sMode:=Q.FindField('E_MODESAISIE').AsString ;
if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q,TypeAction)
                                  else TrouveEtLanceSaisie(Q,taConsult,'N') ;
end;

procedure TFMulSuivACC.ValideAcc ;
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

procedure TFMulSuivACC.BOuvrirClick(Sender: TObject);
Var Ope : String ;
    i   : integer ;
begin
// 14428
if (FListe.NbSelected=0) and (Not FListe.AllSelected) then begin
  HM.Execute(3,caption,''); // Vous n'avez rien sélectionné.
  exit;
end;

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

procedure TFMulSuivACC.BCtrlRibClick(Sender: TObject);
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

procedure TFMulSuivACC.FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
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


procedure TFMulSuivACC.FLettrageChange(Sender: TObject);
begin
  inherited;
if FLettrage.Value='LE' then XX_WHERELETTRAGE.Text:='E_LETTRAGE<>""' else
 if FLettrage.Value='NL' then XX_WHERELETTRAGE.Text:='(E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL")' else
   XX_WHERELETTRAGE.Text:='' ;
end;

procedure TFMulSuivACC.ClickModifRib ;
Var RJal,RExo : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
begin
If ModifRibSurMul(Q,FALSE,TRUE) Then
   BEGIN
   Application.ProcessMessages ; BChercheClick(Nil) ;
   RJal:=Q.FindField('E_JOURNAL').AsString ;
   RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
   RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
   RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
   RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
   RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
   Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
            VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]),[])
   END ;

end;

procedure TFMulSuivACC.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
//Var S : Boolean ;
begin
  inherited;
//S:=(Shift=[ssCtrl]) ;
Case Key of
  VK_F5 : BEGIN Key:=0 ; ClickModifRib ; END
  END ;
end;

procedure TFMulSuivACC.FListeRowEnter(Sender: TObject);
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

procedure TFMulSuivACC.BChercheClick(Sender: TObject);
begin
  inherited;
  FListeRowEnter(nil); // 10930
end;


{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 26/10/2005
Modifié le ... :   /  /    
Description .. : Surcharge de la méthode de Mul.pas pour remplacer le 
Suite ........ : OnDrawDataCell qui ne fonctionne pas
Mots clefs ... : 
*****************************************************************}
procedure TFMulSuivACC.DessineCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
Var sRIB : String ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
if ((Column.FieldName='E_AUXILIAIRE') and (Q.FindField('E_RIB')<>Nil)) then
   BEGIN
   sRIB:=Q.FindField('E_RIB').AsString ; if sRib<>'' then Exit ;
   FListe.Canvas.Brush.Color:= clRed ; FListe.Canvas.Brush.Style:=bsSolid ; FListe.Canvas.Pen.Color:=clRed ;
   FListe.Canvas.Pen.Mode:=pmCopy ; FListe.Canvas.Pen.Width:= 1 ;
   FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
   END ;
end;

{b fb 19/06/2006 FQ12487}
procedure TFMulSuivACC.E_GENERALElipsisClick(Sender: TObject);
begin
  inherited;
  LookUpList(TControl(Sender),'Recherche d''un compte collectif','GENERAUX','G_GENERAL','G_LIBELLE',
  '(G_SUIVITRESO="ENC" OR G_SUIVITRESO="MIX") AND (G_COLLECTIF="X" OR G_LETTRABLE="X")','G_GENERAL',true,-1);
end;

procedure TFMulSuivACC.E_AUXILIAIREElipsisClick(Sender: TObject);
begin
  inherited;
  LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',
  '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")','T_AUXILIAIRE',true,-1)
end;
{e fb 19/06/2006 FQ12487}
end.

