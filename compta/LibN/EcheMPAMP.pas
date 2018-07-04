{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 22/05/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit EcheMPAMP;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Controls,  // mrOK
  Forms,
  StdCtrls,  // TLabel...
  Hctrls,    // THLabel, THValComboBox, OpenSQL, Ferme, AddMenuPop, RechDom, CallHelpTopic, ReadTokenSt
  Mask,      // TMaskEdit
  Buttons,   // TBitBtn
  ExtCtrls,  // TPanel
  Ent1,      // VH
  HEnt1,     // String3, SyncrDefault, taModif, SourisNormale, StrFMontant, V_PGI, IsValidDate
  Saisutil,  // NbJoursOK
  hmsgbox,   // THMsgBox
{$IFDEF EAGLCLIENT}
  UTOB,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  HSysMenu,  // THSystemMenu
  HTB97,     // TToolbarButton97
  FichComm,  // FicheRIB_AGL
  UtilPGI    // EncodeRIB
  ;

Type T_ECHEMPAMP  = RECORD
                  DateEche,DateComptable : TDateTime ;
                  ModePaie,Jal,NatP,CodeAcc  : String3 ;
                  Aux,RIB,RefExterne,RefLibre,NumTraChq : String ;
                  Montant : Double ;
                  NumP : Integer ;
                  TIDTIC : Boolean ;
                  END ;

type
  TFEcheMPAMP = class(TForm)
    PanelBouton: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    E_JOURNAL: TLabel;
    E_NATUREPIECE: TLabel;
    E_DATECOMPTABLE: TLabel;
    E_NUMEROPIECE: TLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    FDateEche2: TMaskEdit;
    FModepaie2: THValComboBox;
    FTitreSup: TLabel;
    HME: THMsgBox;
    HMTrad: THSystemMenu;
    Label5: TLabel;
    E_DATEECHEANCE: TLabel;
    Label7: TLabel;
    E_MONTANT: TLabel;
    Label9: TLabel;
    E_MODEPAIE: TLabel;
    MP_CODEACCEPT: TLabel;
    TMP_CODEACCEPT2: THLabel;
    E_AUXILIAIRE: TLabel;
    T_LIBELLE: TLabel;
    ZoomRib: TToolbarButton97;
    Label6: TLabel;
    TE_REFEXTERNE: THLabel;
    FRefTire: TEdit;
    E_RIB: TEdit;
    TE_REFLIBRE: THLabel;
    FRefTireur: TEdit;
    TE_NUMTRAITECHQ: THLabel;
    E_NUMTRAITECHQ: TEdit;
    MP_CODEACCEPT2: THValComboBox;
    E_RIB2: TEdit;
    E_RIB3: TEdit;
    E_RIB4: TEdit;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BAideClick(Sender: TObject);
    procedure ZoomRibClick(Sender: TObject);
    procedure E_RIBKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure E_RIBDblClick(Sender: TObject);
    procedure FModepaie2Change(Sender: TObject);
  private
  public
    EU : T_ECHEMPAMP ;
  end;

function ModifUneEcheanceMPAMP ( Var EU : T_ECHEMPAMP ) : boolean ;

implementation

{$R *.DFM}

function ModifUneEcheanceMPAMP ( Var EU : T_ECHEMPAMP ) : boolean ;
Var X : TFEcheMPAMP ;
    ii  : integer ;
BEGIN
X:=TFEcheMPAMP.Create(Application) ;
 Try
  X.EU:=EU ;
  ii:=X.ShowModal ;
  Result:=(ii=mrOk) ;
  EU:=X.EU ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

Function ModifLeRIB (Var RIB,Aux : String ) : boolean ;
Var Num{,Nb} : integer ;
    Q      : TQuery ;
    Etab,Guichet,NumCompte,Cle,Dom,Iban,Pays : String ;
BEGIN
Result:=False ;
Num:=FicheRIB_AGL(Aux,taModif,True,RIB,FALSE) ; if Num<=0 then Exit ;
Q:=OpenSQL('Select * from RIB Where R_AUXILIAIRE="'+Aux+'" AND R_NUMERORIB='+IntToStr(Num),True) ;
if Not Q.EOF then
   BEGIN
   Etab:=Q.FindField('R_ETABBQ').AsString ;
   Guichet:=Q.FindField('R_GUICHET').AsString ;
   NumCompte:=Q.FindField('R_NUMEROCOMPTE').AsString ;
   Cle:=Q.FindField('R_CLERIB').AsString ;
   Dom:=Q.FindField('R_DOMICILIATION').AsString ;
   Pays:=Q.FindField('R_PAYS').AsString ;
   Iban:=Q.FindField('R_CODEIBAN').AsString ;
   END ;
Ferme(Q) ;
if (codeisodupays(Pays) <> 'FR') then
  RIB := Iban
else
  RIB:=EncodeRIB(Etab,Guichet,NumCompte,Cle,Dom) ;
Result:=True ;
END ;



procedure TFEcheMPAMP.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TFEcheMPAMP.FormShow(Sender: TObject);
Var Q : TQuery ;
begin
If EU.RIB='' Then EU.RIB:=HME.Mess[2] ;
E_Auxiliaire.Caption:=EU.Aux ; E_RIB.Text:=EU.RIB ; T_Libelle.Caption:='' ;
FRefTire.TExt:=EU.RefExterne ; FRefTireur.TExt:=EU.RefLibre ;
E_NUMTRAITECHQ.Text:=EU.NumTraChq ;
If EU.TIDTIC Then Q:=OpenSQL('Select G_LIBELLE FROM GENERAUX WHERE G_GENERAL="'+Trim(EU.Aux)+'"',TRUE)
             Else Q:=OpenSQL('Select T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+Trim(EU.Aux)+'"',TRUE) ;
If Not Q.Eof Then T_Libelle.Caption:=Q.Fields[0].AsString ;
  Ferme(Q) ;
  SourisNormale ;
FDateEche2.Text:=DateToStr(EU.DateEche) ; FDateEche2.Text:=DateToStr(EU.DateEche) ;
FModePaie2.Value:=EU.ModePaie ; FModePaie2.Value:=EU.ModePaie ;
MP_CODEACCEPT2.Value:=EU.CodeAcc ;
MP_CODEACCEPT.Caption:=MP_CODEACCEPT2.Text ;
E_JOURNAL.Caption:=RechDom('ttJournaux',EU.Jal,False) ; E_NATUREPIECE.Caption:=RechDom('ttNaturePiece',EU.NatP,False) ;
E_NUMEROPIECE.Caption:=IntToStr(EU.NumP) ; E_DATECOMPTABLE.Caption:=DateToStr(EU.DateComptable) ;
E_DATEECHEANCE.Caption:=DateToStr(EU.DateEche) ;
E_MONTANT.Caption:=StrFMontant(EU.Montant,12,V_PGI.OkDecV,'',TRUE) ;
E_MODEPAIE.Caption:=FmodePaie2.Text ;
end;

procedure TFEcheMPAMP.BValiderClick(Sender: TObject);
Var DDE : TDateTime ;
begin
  if Not IsValidDate(FDateEche2.Text) then BEGIN
    HME.Execute(0,'','') ;  // Vous devez renseigner une date valide.
    Exit ;
  END ;
DDE:=StrToDate(FDateEche2.Text) ;
  if Not NbJoursOK(EU.DateComptable,DDE) then BEGIN
    HME.Execute(1,'','') ;  // La date d'échéance doit respecter la plage de saisie autorisée.
    Exit ;
  END ;
EU.DateEche:=DDE ;
EU.ModePaie:=FModePaie2.Value ;
If Trim(E_RIB.Text)=Trim(HME.Mess[2]) Then E_RIB.Text:='' ;

if (Pos('/',E_RIB.Text)=0) and
   {JP 28/10/05 : Problème chez SIC : Si le rib est vide, on évite de se retrouver avec une '*'}
   (E_RIB.Text <> '') then EU.Rib := '*' + E_RIB.Text
                      else EU.Rib := E_RIB.Text ;
EU.RefExterne:=Trim(FRefTire.Text) ; EU.RefLibre:=Trim(FRefTireur.Text) ;
EU.CodeAcc:=MP_CODEACCEPT2.Value ; EU.NumTraChq:=Trim(E_NUMTRAITECHQ.Text) ;
ModalResult:=mrOk ;
end;

procedure TFEcheMPAMP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ((Shift=[]) and (Key=VK_F10)) then
   BEGIN
   Key:=0 ; if BValider.CanFocus then BValider.SetFocus ;
   BValiderClick(Nil) ;
   END ;
end;

procedure TFEcheMPAMP.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFEcheMPAMP.ZoomRibClick(Sender: TObject);
Var Rib,Aux : String ;
begin
Rib:=E_RIB.Text ; Aux:=E_AUXILIAIRE.Caption ;
If ModifLeRIB(Rib,Aux) Then
   BEGIN
    If RIB='' Then Rib := HME.Mess[2] ;  // Non renseigné.
   E_RIB.Text:=Rib ;
   END ;
end;

procedure TFEcheMPAMP.E_RIBKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var Vide : Boolean ;
begin
Vide:=(Shift=[]) ;
if Vide AND (Key=VK_F5) then BEGIN Key:=0 ; ZoomRIBClick(Nil) ; END ;
Key:=0 ;
end;

procedure TFEcheMPAMP.E_RIBDblClick(Sender: TObject);
begin
ZoomRIBClick(Nil) ;
end;

procedure TFEcheMPAMP.FModepaie2Change(Sender: TObject);
Var St,sMode,sCat,Junk : String ;
    i : integer ;
    Okok : boolean ;
begin
OkOk:=FALSE ;
for i:=0 to VH^.MPACC.Count-1 do
    BEGIN
    St:=VH^.MPACC[i] ;
    sMode:=ReadtokenSt(St) ; Junk:=ReadtokenSt(St) ; sCat:=ReadtokenSt(St) ;
    if (sCat='LCR') And (SMode=FModePaie2.Value) then BEGIN OkOk:=TRUE ; Break ; END ;
    END ;
If Not OkOk Then
  BEGIN
  MP_CODEACCEPT2.ItemIndex:=-1 ; MP_CODEACCEPT2.Enabled:=FALSE ; TMP_CODEACCEPT2.Enabled:=FALSE ;
  END Else BEGIN MP_CODEACCEPT2.Enabled:=TRUE ; TMP_CODEACCEPT2.Enabled:=TRUE ; END ;
end;


end.
