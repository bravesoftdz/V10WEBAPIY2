unit Contact;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, DBCGrids, DB, DBTables, Hctrls, DBCtrls,
  Mask, hmsgbox, Ent1, HEnt1, Grids, DBGrids, HSysMenu, MajTable, Hqry;

Procedure FicheContact(Auxiliaire : String ; Mode : TActionFiche) ;

type
  TFContact = class(TForm)
    TContact: THTable;
    SContact : TDataSource;
    FListe   : TDBGrid;
    DBNav    : TDBNavigator;
    Panel1   : TPanel;
    Bevel1   : TBevel;
    FAutoSave: TCheckBox;
    BFirst   : TBitBtn;
    BPrev    : TBitBtn;
    BNext    : TBitBtn;
    BLast    : TBitBtn;
    BDelete  : TBitBtn;
    BInsert  : TBitBtn;
    BImprimer: TBitBtn;
    BValider : TBitBtn;
    BAnnuler : TBitBtn;
    BAide    : TBitBtn;
    BFerme   : TBitBtn;
    MsgBox   : THMsgBox;
    Pappli   : TPanel;
    TC_NOM           : THLabel;
    C_NOM            : TDBEdit;
    TC_SERVICE       : THLabel;
    C_SERVICE        : TDBEdit;
    TC_TELEPHONE     : THLabel;
    C_TELEPHONE      : TDBEdit;
    TC_RVA           : THLabel;
    C_RVA            : TDBEdit;
    TC_FONCTION      : THLabel;
    TC_FAX           : THLabel;
    C_FAX            : TDBEdit;
    TC_CIVILITE      : THLabel;
    C_PRINCIPAL      : TDBCheckBox;
    TC_NUMEROCONTACT : THLabel;
    C_NUMEROCONTACT  : TDBEdit;
    C_TELEX          : TDBEdit;
    TC_TELEX         : THLabel;
    TC_FONCTIONCODEE : THLabel;
    C_CIVILITE       : THDBValComboBox;
    C_FONCTIONCODEE: TDBEdit;
    C_FONCTION: THDBValComboBox;
    TContactC_AUXILIAIRE    : TStringField;
    TContactC_NUMEROCONTACT : TIntegerField;
    TContactC_PRINCIPAL     : TStringField;
    TContactC_NOM           : TStringField;
    TContactC_SERVICE       : TStringField;
    TContactC_FONCTION      : TStringField;
    TContactC_TELEPHONE     : TStringField;
    TContactC_FAX           : TStringField;
    TContactC_TELEX         : TStringField;
    TContactC_SOCIETE       : TStringField;
    TContactC_CIVILITE      : TStringField;
    TContactC_FONCTIONCODEE : TStringField;
    XX_WHERE: TPanel;
    HMTrad: THSystemMenu;
    TContactC_RVA: TStringField;
    procedure FormShow(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SContactDataChange(Sender: TObject; Field: TField);
    procedure SContactStateChange(Sender: TObject);
    procedure SContactUpdateData(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BImprimerClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private    { Déclarations privées }
    Auxiliaire : String ;
    LastNum : Longint ;
    Modifier : Boolean ;
    Mode : TActionFiche ;
    Function Bouge(Button: TNavigateBtn) : boolean ;
    Function EnregOK : boolean ;
    Procedure NewEnreg ;
    Function OnSauve : boolean ;
    Procedure ChargeEnreg ;
    Function  PrincipalUnique : Boolean ;
  public     { Déclarations publiques }
  end;

implementation

{$R *.DFM}
Uses PrintDBG ;

Procedure FicheContact(Auxiliaire : String ; Mode : TActionFiche) ;
var FContact: TFContact;
BEGIN
if Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FContact:=TFContact.Create(Application) ;
 try
  FContact.Auxiliaire:=Auxiliaire ;
  FContact.Mode:=Mode ;
  FContact.ShowModal ;
 Finally
  FContact.free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

Function TFContact.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ; Modifier:=False ;
if TContact.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then Bouge(nbCancel) ;
  mrNo :  if not Bouge(nbCancel) then Exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFContact.EnregOK : boolean ;
var Q : TQuery ;
BEGIN
result:=FALSE  ;
if TContact.state in [dsinsert,dsedit]=False then Exit ;
if TContact.state in [dsinsert,dsedit] then
   BEGIN
   if TContactC_NOM.AsString='' then BEGIN C_NOM.SetFocus ; MsgBox.execute(2,'','') ; Exit ; END ;
   if Not PrincipalUnique then
      BEGIN
      MsgBox.Execute(6,'','') ; C_PRINCIPAL.SetFocus ; Exit ;
      END ;
   if ((TContact.State=dsInsert) or ((TContact.State=dsEdit) and (C_Nom.Modified))) then
     BEGIN
     Q:=OpenSQL('SELECT C_NOM FROM CONTACT WHERE C_AUXILIAIRE="'+AUXILIAIRE+'" AND C_NOM="'+C_NOM.Text+'"',true);
     if (Not Q.EOF) then
       BEGIN
       Case MsgBox.Execute(4,'','') of
         mrYes    : ;
         mrNo     : BEGIN Bouge(nbCancel) ; Ferme(Q) ; Exit ; END ;
         mrCancel :BEGIN C_NOM.SetFocus ; Ferme(Q) ; Exit ; END ;
         END ;
       END ;
     Ferme(Q) ;
     END ;
   END else Exit ;
result:=TRUE  ;
END ;

Procedure TFContact.ChargeEnreg ;
var Q : TQuery ;
    Lib : String ;
BEGIN
Q:=OpenSQL('Select T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+AUXILIAIRE+'"',True) ;
if not Q.Eof then Lib:=Q.Fields[0].AsString else Lib:='' ;
Ferme(Q) ;
InitCaption(Self,AUXILIAIRE+' '+Lib,'') ;
if Mode=taConsult then BEGIN FicheReadOnly(Self) ; Exit ; END ;
END ;

Procedure TFContact.NewEnreg ;
BEGIN
InitNew(TContact) ;
TContactC_AUXILIAIRE.AsString:=Auxiliaire ;
TContactC_NumeroContact.AsInteger:=LastNum+1 ;
C_NOM.SetFocus ;
END ;

Function TFContact.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
Result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,nbfirst,nbinsert :
     BEGIN
     if Not OnSauve then Exit ;
     END ;
   nbPost : if Not EnregOK then exit ;
   nbDelete : if MsgBox.execute(1,'','')<>mrYes then exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(Msgbox.Mess[5]);
Result:=TRUE ;
if Button=NbInsert then
   BEGIN
   DBNav.BtnClick(Nblast) ; LastNum:=TContactC_NumeroContact.AsInteger ;
   DBNav.BtnClick(Button) ; NewEnreg ;
   END ;
END ;

procedure TFContact.FormShow(Sender: TObject);
begin
TContact.Open ;
if Auxiliaire<>'' then TContact.SetRange([Auxiliaire,0],[Auxiliaire,99999]) ;
Case Mode Of
     taConsult : FicheReadOnly(Self) ;
     taCreat..taCreatOne : Bouge(nbInsert) ;
     else if (TContact.Bof) and (TContact.Eof) then Bouge(nbInsert) ;
   End ;
end;

procedure TFContact.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFContact.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFContact.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFContact.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFContact.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFContact.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFContact.BValiderClick(Sender: TObject);
begin
Modifier:=False ;
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then Bouge(nbInsert) ;
   if Mode=taCreatOne then Close ;
   END ;
end;

procedure TFContact.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFContact.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFContact.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFContact.SContactDataChange(Sender: TObject; Field: TField);
var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not TContact.Modified ; BDelete.Enabled:=(Not TContact.Modified) ;
if(TContact.Eof)And(TContact.Bof) then BDelete.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TContact.BOF;
   DnEnable := Enabled and not TContact.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
//  code de liaison entre field
//   if ((Field.FieldName='T_LIBELLE') and (T_ABREGE.Field.AsString='')) then
//      T_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFContact.SContactStateChange(Sender: TObject);
begin  Modifier:=True ; end;

procedure TFContact.SContactUpdateData(Sender: TObject);
begin
if Modifier then BEGIN Modifier:=False ; OnSauve ; END ;
end;

Function TFContact.PrincipalUnique : Boolean ;
Var Q : TQuery ;
BEGIN
Result:=True ;
if Not C_PRINCIPAL.Checked then Exit ;
Q:=OpenSql('Select C_PRINCIPAL from CONTACT Where C_AUXILIAIRE="'+Auxiliaire+'" And '+
           'C_PRINCIPAL="X" And C_NUMEROCONTACT<>'+IntToStr(TContactC_NUMEROCONTACT.AsInteger),True) ;
if Not Q.Eof then Result:=False ;
Ferme(Q) ;
END ;

procedure TFContact.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if((ssCtrl in Shift)AND(Key=VK_DELETE)AND(TContact.Eof)) then BEGIN Key:=0 ; Exit ; END ;
if((ssCtrl in Shift)AND(Key=VK_DELETE))then Key:=0 ; 
end;

procedure TFContact.BImprimerClick(Sender: TObject);
begin
XX_WHERE.Hint:=' WHERE C_AUXILIAIRE="'+Auxiliaire+'"' ;
PrintDBGrid(Nil ,XX_WHERE,Caption,'PRT_CONTACT') ;
end;

procedure TFContact.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
