unit SaisTVA;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Buttons, Ent1, ExtCtrls, HSysMenu, HEnt1
  , uLibPieceCompta
  , uTOB
  , SaisUtil         // ChangeFormatDevise
  ;

procedure CPLanceFiche_ControleTva( vPiece : TPieceCompta ; vNumLigne : Integer ) ;

type
  TFSaisTVA = class(TForm)
    GroupBox1: TGroupBox;
    HTPF1: THLabel;
    HLabel12: THLabel;
    HLabel1: THLabel;
    SGP: THNumEdit;
    SAP: THNumEdit;
    SFP: THNumEdit;
    SFL: THNumEdit;
    SAL: THNumEdit;
    SGL: THNumEdit;
    SGC: TEdit;
    SAC: TEdit;
    SFC: TEdit;
    GroupBox2: TGroupBox;
    HTPF2: THLabel;
    HLabel7: THLabel;
    CAP: THNumEdit;
    CFP: THNumEdit;
    CFL: THNumEdit;
    CAL: THNumEdit;
    CAC: TEdit;
    CFC: TEdit;
    GroupBox3: TGroupBox;
    HTPF3: THLabel;
    HLabel17: THLabel;
    EAP: THNumEdit;
    EFP: THNumEdit;
    EFL: THNumEdit;
    EAL: THNumEdit;
    EAC: TEdit;
    EFC: TEdit;
    BValide: THBitBtn;
    BAide: THBitBtn;
    HLabel15: THLabel;
    HLabel14: THLabel;
    HLabel11: THLabel;
    Panel1: TPanel;
    HINFO: TLabel;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure ChargeInfos ;
  private
  public
    Decim : integer ;
    FPiece  : TPieceCompta ;
    FInIdx : Integer ;
  end;

implementation

uses
    {$IFDEF MODENT1}
    CPProcMetier,
    {$ENDIF MODENT1}
    UtilPGI
    {$IFNDEF EAGLCLIENT}
    ,PrintDBG
    {$ENDIF}
    ; //XVI 24/02/2005

{$R *.DFM}

procedure TFSaisTVA.FormShow(Sender: TObject);
Var i : integer ;
begin
  ChargeInfos ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
for i:=0 to ComponentCount-1 do if Components[i] is THNumedit
    then THNumEdit(Components[i]).Value:=Arrondi(THNumEdit(Components[i]).Value,Decim) ;
if EAP.Value<>0 then EAP.Font.Color:=clRed ; if EFP.Value<>0 then EFP.Font.Color:=clRed ;
if EAL.Value<>0 then EAL.Font.Color:=clRed ; if EFL.Value<>0 then EFL.Font.Color:=clRed ;
{$IFDEF CCS3}
if VH^.PaysLocalisation<>CodeISOES then //XVI 24/02/2005
 begin
  SFC.Visible:=False ; SFL.Visible:=False ; SFP.Visible:=False ; HTPF1.Visible:=False ;
  CFC.Visible:=False ; CFL.Visible:=False ; CFP.Visible:=False ; HTPF2.Visible:=False ;
  EFC.Visible:=False ; EFL.Visible:=False ; EFP.Visible:=False ; HTPF3.Visible:=False ;
 End ;
{$ENDIF ESP}
end;

procedure TFSaisTVA.BValideClick(Sender: TObject);
begin
Close ;
end;

procedure TFSaisTVA.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;


//===================================================
procedure CPLanceFiche_ControleTVA( vPiece : TPieceCompta ; vNumLigne : Integer ) ;
var X : TFSaisTVA ;
begin
  // Appel de la form VisuTVA
  X := TFSaisTVA.Create(Application) ;
  try
    X.Decim  := vPiece.Devise.Decimale ;
    X.FPiece := vPiece ;
    X.FInIdx := vNumLigne ;
    X.ShowModal ;
  finally
    X.Free ;
  end ;
end ;
//===================================================

procedure TFSaisTVA.ChargeInfos;
Var //LigTVA           : Integer ;
    //LigTPF           : Integer ;
    // Montants saisis dans les lignes
    lDebitHT         : Double ;
    lCreditHT        : Double ;
    lTotalDebitHT    : Double ;
    lTotalCreditHT   : Double ;
    lDebitTVA        : Double ;
    lCreditTVA       : Double ;
    lTotalDebitTVA   : Double ;
    lTotalCreditTVA  : Double ;
    lDebitTPF        : Double ;
    lCreditTPF       : Double ;
    lTotalDebitTPF   : Double ;
    lTotalCreditTPF  : Double ;
    // Montants calculés
    lCalcDebitTVA        : Double ;
    lCalcCreditTVA       : Double ;
    lCalcTotalDebitTVA   : Double ;
    lCalcTotalCreditTVA  : Double ;
    lCalcDebitTPF        : Double ;
    lCalcCreditTPF       : Double ;
    lCalcTotalDebitTPF   : Double ;
    lCalcTotalCreditTPF  : Double ;
    // Autres infos
    lStCptHT         : String ;
    lStCptTVA        : String ;
    lStCptTPF        : String ;
//    lTauxTva         : Double ;
//    lTauxTpf         : Double ;
    lBoAchat         : Boolean ;
//    lBoSurEnc        : boolean ;
    lBoSoumisTPF     : boolean ;
    lTobEcr          : TOB ;
    lTobTva          : TOB ;
    lTobTpf          : TOB ;
    lStCodeTva       : String ;
    lStCodeTpf       : String ;
    lStRegimeTva     : String ;
begin
  // Init des variables
  lDebitHT         := 0 ;
  lCreditHT        := 0 ;
  lTotalDebitHT    := 0 ;
  lTotalCreditHT   := 0 ;
  lDebitTVA        := 0 ;
  lCreditTVA       := 0 ;
  lTotalDebitTVA   := 0 ;
  lTotalCreditTVA  := 0 ;
  lDebitTPF        := 0 ;
  lCreditTPF       := 0 ;
  lTotalDebitTPF   := 0 ;
  lTotalCreditTPF  := 0 ;

  // Recup infos régime, Codes TVA et TPF du compte HT
  lTobEcr := FPiece.GetTob( FInIdx ) ;
  if lTobEcr=Nil then Exit ;

  lStCodeTva   := lTobEcr.GetString('E_TVA') ;
  lStCodeTpf   := lTobEcr.GetString('E_TPF') ;
//  lBoSurEnc    := lTobEcr.GetValue('E_TVAENCAISSEMENT') = 'X' ;
  lBoAchat     := FPiece.EstAchat( FInIdx ) ;
  lStRegimeTva := FPiece.GetRegimeTVA( FInIdx ) ;
  lBoSoumisTPF := FPiece.EstSoumisTPF( FInIdx ) ;

    //  AttribRegimeEtTva ;
//  lTauxTva    := Tva2Taux( lStRegimeTva, lStCodeTva, FPiece.EstAchat ) ;
//  lTauxTpf    := Tpf2Taux( lStRegimeTva, lStCodeTpf, FPiece.EstAchat ) ;

  lStCptTVA := FPiece.GetCompteTva( FInIdx ) ;
  lStCptTPF := FPiece.GetCompteTpf( FInIdx ) ;

  // Calculs et recup de la ligne HT
  FPiece.GetInfosLigne( FInIdx, lDebitHT, lCreditHT, lTotalDebitHT, lTotalCreditHT ) ;
  lStCptHT := FPiece.GetValue( FInIdx, 'E_GENERAL' ) ;

  // Calculs et recup de la ligne TVA
  if lStCptTVA<>'' then
    begin
    lTobTva := FPiece.TrouveLigneCompte( lStCptTva ) ;
    if lTobTva <> nil then
      FPiece.GetInfosLigne( lTobTva.GetValue('E_NUMLIGNE'), lDebitTva, lCreditTVA, lTotalDebitTVA, lTotalCreditTVA ) ;
    end ;

  lCalcDebitTVA       := HT2TVA( lDebitHT,       lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, Decim ) ;
  lCalcCreditTVA      := HT2TVA( lCreditHT,      lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, Decim ) ;
  lCalcTotalDebitTVA  := HT2TVA( lTotalDebitHT,  lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, Decim ) ;
  lCalcTotalCreditTVA := HT2TVA( lTotalCreditHT, lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, Decim ) ;


  // Calculs et recup de la ligne TPF
  if lStCptTPF<>'' then
    BEGIN
    lTobTPF := FPiece.TrouveLigneCompte( lStCptTPF ) ;
    if lTobTPF <> nil then
      FPiece.GetInfosLigne( lTobTPF.GetValue('E_NUMLIGNE'), lDebitTPF, lCreditTPF, lTotalDebitTPF, lTotalCreditTPF ) ;
    END ;

  lCalcDebitTPF       := HT2TPF( lDebitHT,       lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, Decim ) ;
  lCalcCreditTPF      := HT2TPF( lCreditHT,      lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, Decim ) ;
  lCalcTotalDebitTPF  := HT2TPF( lTotalDebitHT,  lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, Decim ) ;
  lCalcTotalCreditTPF := HT2TPF( lTotalCreditHT, lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, Decim ) ;

  ChangeFormatDevise( self, Decim, FPiece.Devise.Symbole) ;

  // ======================
  // ==== MAJ INTERFACE ===
  // ======================

  // Comptes
  SGC.Text := lStCptHT  ;
  SAC.Text := lStCptTVA ;
  CAC.Text := lStCptTVA ;
  EAC.Text := lStCptTVA ;
  SFC.Text := lStCptTPF ;
  CFC.Text := lStCptTPF ;
  EFC.Text := lStCptTPF ;

  // Saisie
  AfficheLeSolde( SGL, lDebitHT,       lCreditHT ) ;
  AfficheLeSolde( SGP, lTotalDebitHT,  lTotalCreditHT ) ;
  AfficheLeSolde( SAL, lDebitTVA,      lCreditTVA ) ;
  AfficheLeSolde( SAP, lTotalDebitTVA, lTotalCreditTVA ) ;
  AfficheLeSolde( SFL, lDebitTPF,      lCreditTPF ) ;
  AfficheLeSolde( SFP, lTotalDebitTPF, lTotalCreditTPF ) ;

  // Calculs
  AfficheLeSolde( CAL, lCalcDebitTVA,      lCalcCreditTVA ) ;
  AfficheLeSolde( CAP, lCalcTotalDebitTVA, lCalcTotalCreditTVA ) ;
  AfficheLeSolde( CFL, lCalcDebitTPF,      lCalcCreditTPF ) ;
  AfficheLeSolde( CFP, lCalcTotalDebitTPF, lCalcTotalCreditTPF ) ;

  // Ecarts
  AfficheLeSolde( EAL, ( lDebitTVA - lCalcDebitTVA ),           ( lCreditTVA - lCalcCreditTVA ) ) ;
  AfficheLeSolde( EAP, ( lTotalDebitTVA - lCalcTotalDebitTVA ), ( lTotalCreditTVA - lCalcTotalCreditTVA ) ) ;
  AfficheLeSolde( EFL, ( lDebitTPF - lCalcDebitTPF) ,           ( lCreditTPF - lCalcCreditTPF ) ) ;
  AfficheLeSolde( EFP, ( lTotalDebitTPF - lCalcTotalDebitTPF) , ( lTotalCreditTPF - lCalcTotalCreditTPF ) ) ;

  // Info Ligne
  HInfo.Caption := HInfo.Caption + IntToStr( FInIdx ) + '   '
                                 + FPiece.GetValue( FInIdx , 'E_GENERAL' ) + '   '
                                 + FPiece.GetValue( FInIdx , 'E_REFINTERNE' ) ;

end;

end.
