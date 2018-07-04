unit ParaClo ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Buttons, ExtCtrls, hmsgbox, HSysMenu;


Type ttContrepartie = (Aucune,ChaqueLigne,EnPiedDC,EnPiedSolde) ;
     ttModeGenerePiece = (UneParCpt,UneSeule) ;
     ttCompteBanque = (BqUneSeuleRef,BqUneParRef) ;
     ttComptePointableVentilable = (SurAnalytique,UneSeuleRef,UneParRef) ;

Type tParamcloture = Record
                     CloContrep,AnoContrep : ttContrepartie ;
                     CloPiece,AnoPiece : ttModeGenerePiece ;
                     AnoCompteBanque : ttCompteBanque ;
                     AnoComptePV : ttComptePointableVentilable ;
                     End ;

Function ParametrageCloture(Definitive : Boolean ; Var ParamCloture : TParamCloture ; ExisteCptPV : Boolean ; vBoAnoDyna : boolean ) : Boolean ;

type
  TParamClo = class(TForm)
    GParamClo: TPanel;
    GroupBox3: TGroupBox;
    HLabel7: THLabel;
    HCloPiece: THLabel;
    CloContrep: THValComboBox;
    CloPiece: THValComboBox;
    GroupBox4: TGroupBox;
    HLabel11: THLabel;
    HAnoPiece: THLabel;
    AnoContrep: THValComboBox;
    AnoPiece: THValComboBox;
    HMess: THMsgBox;
    GBPVBQE: TGroupBox;
    ANOCPTBQ: TRadioGroup;
    ANOCptPV: TRadioGroup;
    Panel1: TPanel;
    BCValide: THBitBtn;
    BCAide: THBitBtn;
    BFerme: THBitBtn;
    HMTrad: THSystemMenu;
    procedure BCValideClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure AnoPieceChange(Sender: TObject);
    procedure CloPieceChange(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BCAideClick(Sender: TObject);
  public
    { Déclarations privées }
    ExisteCptPV : Boolean ;
    OnSortSansValider : Boolean ;
    ParaClo : TParamCloture ;
    CloDef : Boolean ;
    AnoDyna : boolean ;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPVersion;
  {$ELSE}
  Ent1;
  {$ENDIF MODENT1}


Function ParametrageCloture(Definitive : Boolean ; Var ParamCloture : TParamCloture ; ExisteCptPV : Boolean ; vBoAnoDyna : boolean ) : Boolean ;
var ParamClo: TParamClo;
begin
ParamClo:=TParamClo.Create(Application) ; ParamClo.OnSortSansValider:=FALSE ; Result:=TRUE ;
 try
   ParamClo.ParaClo:=ParamCloture ;
   ParamClo.CloDef:=Definitive ;
   ParamClo.ExisteCptPV:=ExisteCptPV ;
   ParamClo.AnoDyna := vBoAnoDyna ;
   ParamClo.ShowModal ;
 Finally
   ParamCloture:=ParamClo.ParaClo ;
   If ParamClo.OnSortSansValider Then Result:=FALSE ;
   ParamClo.free ;
 End ;
end ;

procedure TParamClo.BCValideClick(Sender: TObject);
begin
If CloContrep.Value='PID' Then ParaClo.CloContrep:=EnPiedDC Else
   If CloContrep.Value='PIS' Then ParaClo.CloContrep:=EnPiedSolde Else
      If CloContrep.Value='LIG' Then ParaClo.CloContrep:=ChaqueLigne ;
If CloPiece.Value='MON' Then ParaClo.CloPiece:=UneParCpt Else
   If CloPiece.Value='MUL' Then ParaClo.CloPiece:=UneSeule ;
If ANOContrep.Value='PID' Then ParaClo.ANOContrep:=EnPiedDC Else
   If ANOContrep.Value='PIS' Then ParaClo.ANOContrep:=EnPiedSolde Else
      If ANOContrep.Value='LIG' Then ParaClo.ANOContrep:=ChaqueLigne ;
If ANOPiece.Value='MON' Then ParaClo.ANOPiece:=UneParCpt Else
   If ANOPiece.Value='MUL' Then ParaClo.ANOPiece:=UneSeule ;
Case ANOCptBQ.ItemIndex Of
  0 : ParaClo.AnoCompteBanque:=BQUneSeuleRef ;
  1 : ParaClo.AnoCompteBanque:=BQUneParRef ;
  End ;
Case ANOCptPV.ItemIndex Of
  0 : ParaClo.AnoComptePV:=SurAnalytique ;
  1 : ParaClo.AnoComptePV:=UneParRef ;
  2 : ParaClo.AnoComptePV:=UneSeuleRef ;
  END ;
end;

procedure TParamClo.FormShow(Sender: TObject);
begin
  {JP 20/11/07 : FQ 21247 : L'ancien mécanisme d'a-nouveaux par référence est soumis à code spécifique}
  if EstSpecif('51216') then begin
    GBPVBQE.Visible := True;
    Height := 400;
  end
  else begin
    GBPVBQE.Visible := False;
    Height := 240;
  end;

If Not CloDef Then HelpContext:=7742100 ;
Case ParaClo.CloContrep Of
  EnPiedDC : CloContrep.Value:='PID' ;
  EnPiedSolde : CloContrep.Value:='PIS' ;
  ChaqueLigne : CloContrep.Value:='LIG' ;
  END ;
Case ParaClo.ANOContrep Of
  EnPiedDC : ANOContrep.Value:='PID' ;
  EnPiedSolde : ANOContrep.Value:='PIS' ;
  ChaqueLigne : ANOContrep.Value:='LIG' ;
  END ;
Case ParaClo.CloPiece Of
  UneParCpt : CloPiece.Value:='MON' ;
  UneSeule : CloPiece.Value:='MUL' ;
  END ;
Case ParaClo.ANOPiece Of
  UneParCpt : ANOPiece.Value:='MON' ;
  UneSeule : ANOPiece.Value:='MUL' ;
  END ;
Case ParaClo.AnoCompteBanque Of
  BQUneSeuleRef : ANOCptBQ.ItemIndex:=0 ;
  BQUneParRef : ANOCptBQ.ItemIndex:=1 ;
  End ;
Case ParaClo.AnoComptePV Of
  SurAnalytique : ANOCptPV.ItemIndex:=0 ;
  UneParRef : ANOCptPV.ItemIndex:=1 ;
  UneSeuleRef : ANOCptPV.ItemIndex:=2 ;
  END ;
If Not CloDef Then
   BEGIN
   ANOContrep.SetFocus ;
   GroupBox3.Enabled:=FALSE  ;
   Caption:=HMess.Mess[0] ;
   CloContrep.Enabled:=FALSE ;
   CloPiece.Enabled:=FALSE ;
   HLabel7.Enabled:=FALSE ;
   HCloPiece.Enabled:=FALSE ;
   END Else
   BEGIN
   CloContrep.SetFocus ;
   Caption:=HMess.Mess[1] ;
   END ;
ANOCptPV.Enabled:=ExisteCptPV ;
If Not CloDef Then GBPVBQE.Enabled:=FALSE ;
GroupBox4.Enabled := not AnoDyna ;
UpdateCaption(Self) ;
end;

procedure TParamClo.AnoPieceChange(Sender: TObject);
begin
If ANOPiece.Value='MON' Then ANOContrep.Value:='LIG' ;
end;

procedure TParamClo.CloPieceChange(Sender: TObject);
begin
If CloPiece.Value='MON' Then CloContrep.Value:='LIG' ;
end;

procedure TParamClo.BFermeClick(Sender: TObject);
begin
OnSortSansValider:=TRUE ;
end;

procedure TParamClo.BCAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.

