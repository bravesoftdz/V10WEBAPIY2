{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 21/05/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit EcheMPA;

interface

uses
  Windows, Messages, SysUtils, Classes,
  Graphics,
  Controls,
  Forms,
  Mask,
  StdCtrls,
  Hctrls,    // THPanel, THValComboBox, OpenSQL, Ferme, AddMenuPop, RechDom, CallHelpTopic
  Buttons,   // TBitBtn
  ExtCtrls,
  HEnt1,     // String3, SyncrDefault, taModif, SourisNormale, StrFMontant, V_PGI, IsValidDate
  Saisutil,  // NbJoursOK
  hmsgbox,   // THMsgBox
{$IFDEF EAGLCLIENT}
  uTOB,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  HSysMenu,
  HTB97,
  FichComm,  // FicheRIB_AGL
  UtilPGI    // EncodeRIB
  ;

Type T_ECHEMPA  = RECORD
                  DateEche,DateComptable : TDateTime ;
                  ModePaie,Jal,NatP  : String3 ;
                  Aux,RIB,RefExterne : String ;
                  Montant : Double ;
                  NumP : Integer ;
                  END ;

type
  TFEcheMPA = class(TForm)
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
    HLabel1: THLabel;
    E_AUXILIAIRE: TLabel;
    T_LIBELLE: TLabel;
    ZoomRib: TToolbarButton97;
    Label6: TLabel;
    TE_REFEXTERNE: THLabel;
    FRefTire: TEdit;
    E_RIB: TEdit;
    MP_CODEACCEPT2: TEdit;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BAideClick(Sender: TObject);
    procedure FModepaie2Change(Sender: TObject);
    procedure ZoomRibClick(Sender: TObject);
    procedure E_RIBKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure E_RIBDblClick(Sender: TObject);
  private
  public
    EU : T_ECHEMPA ;
  end;

function  ModifUneEcheanceMPA ( Var EU : T_ECHEMPA ) : boolean ;
Function  ModifLeRIB (Var RIB,Aux : String ) : boolean ;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibEcriture, Ent1 ;

function ModifUneEcheanceMPA ( Var EU : T_ECHEMPA ) : boolean ;
Var X : TFEcheMPA ;
    ii  : integer ;
BEGIN
  X := TFEcheMPA.Create(Application) ;
  Try
    X.EU := EU ;
    ii := X.ShowModal ;
    Result := (ii=mrOk) ;
    EU := X.EU ;
  Finally
    X.Free ;
  End ;
  Screen.Cursor := SyncrDefault ;
END ;

Function ModifLeRIB (Var RIB,Aux : String ) : boolean ;
Var Num    : integer ;
    Q      : TQuery ;
    Etab,Guichet,NumCompte,Cle,Dom : String ;
BEGIN
  Result := False ;
  Num := FicheRIB_AGL(Aux,taModif,True,RIB,FALSE) ; if Num<=0 then Exit ;
  Q := OpenSQL('Select * from RIB Where R_AUXILIAIRE="'+Aux+'" AND R_NUMERORIB='+IntToStr(Num),True) ;
  if VH^.PaysLocalisation=CodeISOES then Begin
       Rib:=IBANtoE_RIB(Q.FindField('R_CODEIBAN').AsString) ;
       Result:=(trim(Rib)<>'') ;
  End else
  Begin
    if Not Q.EOF then BEGIN
      if Q.FindField('R_PAYS').asString= 'FRA' then begin {b Lek 070606 FQ12051}
        Etab := Q.FindField('R_ETABBQ').AsString ;
        Guichet := Q.FindField('R_GUICHET').AsString ;
        NumCompte := Q.FindField('R_NUMEROCOMPTE').AsString ;
        Cle := Q.FindField('R_CLERIB').AsString ;
        Dom := Q.FindField('R_DOMICILIATION').AsString ;
        RIB := EncodeRIB(Etab,Guichet,NumCompte,Cle,Dom) ;
        Result := True ;
      END else
      begin
        Rib:=IBANtoE_RIB(Q.FindField('R_CODEIBAN').AsString) ;
        Result:=(trim(Rib)<>'') ;
      end;
    end; {e Lek 070606 FQ12051}
  End ;
  Ferme(Q) ; //XVI 24/02/2005
END ;

procedure TFEcheMPA.FormCreate(Sender: TObject);
begin
  PopUpMenu := ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TFEcheMPA.FormShow(Sender: TObject);
Var Q : TQuery ;
begin
  If EU.RIB='' Then EU.RIB := HME.Mess[2] ; // Non renseigné.
  E_Auxiliaire.Caption := EU.Aux ;
  E_RIB.Text := EU.RIB ;
  T_Libelle.Caption := '' ;
  FRefTire.TExt := EU.RefExterne ;

  Q := OpenSQL('Select T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+Trim(EU.Aux)+'"',TRUE) ;
  If Not Q.Eof Then T_Libelle.Caption := Q.Fields[0].AsString ;
  Ferme(Q) ;

  SourisNormale ;
  FDateEche2.Text := DateToStr(EU.DateEche) ;
  FDateEche2.Text := DateToStr(EU.DateEche) ;
  FModePaie2.Value := EU.ModePaie ;
  FModePaie2.Value := EU.ModePaie ;
  MP_CODEACCEPT.Caption := MP_CODEACCEPT2.Text ;
  E_JOURNAL.Caption := RechDom('ttJournal',EU.Jal,False) ;
  E_NATUREPIECE.Caption := RechDom('ttNaturePiece',EU.NatP,False) ;
  E_NUMEROPIECE.Caption := IntToStr(EU.NumP) ;
  E_DATECOMPTABLE.Caption := DateToStr(EU.DateComptable) ;
  E_DATEECHEANCE.Caption := DateToStr(EU.DateEche) ;
  E_MONTANT.Caption := StrFMontant(EU.Montant,12,V_PGI.OkDecV,'',TRUE) ;
end;

procedure TFEcheMPA.BValiderClick(Sender: TObject);
Var DDE : TDateTime ;
begin
  if Not IsValidDate(FDateEche2.Text) then BEGIN
    HME.Execute(0,'','') ; // Vous devez renseigner une date valide.
    Exit ;
  END ;
  DDE := StrToDate(FDateEche2.Text) ;
  if Not NbJoursOK(EU.DateComptable,DDE) then BEGIN
    HME.Execute(1,'','') ; // La date d'échéance doit respecter la plage de saisie autorisée.
    Exit ;
  END ;
  EU.DateEche := DDE ;
  EU.ModePaie := FModePaie2.Value ;
  If Trim(E_RIB.Text)=Trim(HME.Mess[2]) Then E_RIB.Text := '' ; // Non renseigné.
  EU.Rib := E_RIB.Text ;
  EU.RefExterne := Trim(FRefTire.Text) ;
  ModalResult := mrOk ;
end;

procedure TFEcheMPA.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((Shift=[]) and (Key=VK_F10)) then BEGIN
    Key := 0 ;
    if BValider.CanFocus then BValider.SetFocus ;
    BValiderClick(Nil) ;
  END ;
end;

procedure TFEcheMPA.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end;

procedure TFEcheMPA.FModepaie2Change(Sender: TObject);
Var Q : TQuery ;
    CodeAcc : String ;
begin
  CodeAcc := '' ;
  Q := OpenSQL('SELECT MP_CODEACCEPT, MP_CATEGORIE FROM MODEPAIE WHERE MP_MODEPAIE="'+FModePaie2.Value+'"',TRUE) ;
  If Not Q.Eof Then BEGIN
    If Q.Fields[1].AsString='LCR' Then CodeAcc := RechDom('TTLETTRECHANGE',Q.Fields[0].AsString,FALSE) ;
  END ;
  If CodeAcc='Error' Then CodeAcc := '' ;
  MP_CODEACCEPT2.Text := CodeAcc ;
  Ferme(Q) ;
  SourisNormale ;
end;

procedure TFEcheMPA.ZoomRibClick(Sender: TObject);
Var Rib,Aux : String ;
begin
  Rib := E_RIB.Text ;
  Aux := E_AUXILIAIRE.Caption ;
  If ModifLeRIB(Rib,Aux) Then BEGIN
    If RIB='' Then Rib := HME.Mess[2] ; // Non renseigné.
    E_RIB.Text := Rib ;
  END ;
end;

procedure TFEcheMPA.E_RIBKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var
  Vide : Boolean ;
begin
  Vide := (Shift=[]) ;
  if Vide AND (Key=VK_F5) then BEGIN
    Key := 0 ;
    ZoomRIBClick(Nil) ;
  END ;
  Key := 0 ;
end;

procedure TFEcheMPA.E_RIBDblClick(Sender: TObject);
begin
  ZoomRIBClick(Nil) ;
end;

end.
