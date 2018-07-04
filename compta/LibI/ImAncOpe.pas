// FQ 17339 - MBO - 20/01/2006 - si date fin contrat modifiée lors de la levée option
//                               affichage de la zone i_datesuspcb en date fin contrat
// FQ 17446 - BTY - 03/06 - Passer l'information i_journala à la modif de bases
// FQ 18395 - MBO - 15/06/2006 - recupération de la date début d'amortissement économique
// FQ 18393 - MBO - 06/06 - En série, reprendre la date de l'opération
// BTY - 10/06 - reprendre le montant de la prime d'équipement
// BTY - 11/06 - reprendre le montant de la subvention d'investissement
// BTY - 04/07 - reprendre la DPI
// MBO - 12/06/07 - Rev9 sur modif de base / Modif de la variable fremplace (lecture de i_string1)
// BTY 07/07 FQ 17339 suite : 01/01/1900 remplacé par stDate1900 norme internationale

unit ImAncOpe;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  ComCtrls,
  HRichEdt,
  HRichOLE,
  Hctrls,
  Mask,
  HTB97,
  ExtCtrls,
  HPanel,
  {$IFDEF EAGLCLIENT}
  uTOB,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  Hent1,
  HSysMenu,
  hmsgbox,
  ParamDat,
  ImContra, (*UtilSoc,*)
  ImEnt;

type
  TFAncOpe = class(TForm)
    HPanel3: THPanel;
    bValider: TToolbarButton97;
    bFermer: TToolbarButton97;
    ToolbarButton976: TToolbarButton97;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    HPanel2: THPanel;
    GroupBox1: TGroupBox;
    tI_IMMO: THLabel;
    laI_IMMO: THLabel;
    tI_LIBELLE: THLabel;
    laI_LIBELLE: THLabel;
    laI_MONTANTHT: THLabel;
    tI_MONTANTHT: THLabel;
    laI_DATEPIECEA: THLabel;
    tI_DATEPIECEA: THLabel;
    HLabel1: THLabel;
    laI_COMPTEIMMO: THLabel;
    HLabel15: THLabel;
    laI_QUANTITE: THLabel;
    GroupBox2: TGroupBox;
    HLabel3: THLabel;
    DATEOPE: THCritMaskEdit;
    PAGES: TPageControl;
    PBLOCNOTE: TTabSheet;
    IL_BLOCNOTE: THRichEditOLE;
    procedure FormShow(Sender: TObject);
    procedure bValiderClick(Sender: TObject);
    procedure bFermerClick(Sender: TObject);
    procedure DATEOPEKeyPress(Sender: TObject; var Key: Char);
    procedure DATEOPEExit(Sender: TObject);
    procedure ToolbarButton976Click(Sender: TObject);
    procedure InitZones;virtual;
    function ControleZones : boolean;virtual;
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Déclarations privées }
    procedure EnregistreOperation (PP : TProc);
  public
    { Déclarations publiques }
    fCode, fLibelle, fNature : string;
    fDateAchat, fDateAmort : TDateTime;
    fMontantHT, fBaseEco, fBaseFisc, fBaseTP, fTauxEco, fMttHT : double;
    fResiduel, fQuantite : double;
    fDureeEco : integer;
    fMethodeEco : string;
    fCompteImmo, fCompteCharge : string;
    fProcEnreg : procedure of object;
    fPlanActif : integer;
    fOrdre, fOrdreS : integer;
    bFiscal : boolean;
    fSaveDate : TDateTime ;
    fContrat : TImContrat;
    fDateFinContrat : TDateTime;
    fTVARecuperable : double;
    fTVARecuperee : double;
    fEtat : string;
    fJournala : string; // BTY 03/06 FQ 17446
    fDateDebEco : TDateTime;  // mbo fq 18395
    fDateSerie : TDateTime;  // BTY FQ 18393
    fSBVPRI : double;   // BTY 10/06
    fSBVMT : double;    // BTY 11/06
    fbFuturVNF : boolean;
    fbDPI  : boolean;
    fbGestionfiscale : boolean;
    fRemplace : string;
    fRepriseDR, fRepriseFEC : double;
    fDateDebFis : TDateTime; 
  end;

implementation

uses Outils{$IFDEF SERIE1} {$ELSE},UtilPGI{$ENDIF};

{$R *.DFM}

procedure TFAncOpe.FormShow(Sender: TObject);
begin
if ImBlocage(['nrCloture'],False,'nrAucun') then Close ;
InitZones;
end;

procedure TFAncOpe.InitZones;
var Q : TQuery;
    fDateLevOpt : TDateTime;
  begin
  Q := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+fCode+'"',True);
  if not Q.Eof then
  begin
    fDateAchat := Q.FindField ('I_DATEPIECEA').AsDateTime;
    fDateAmort:= Q.FindField('I_DATEAMORT').AsDateTime ;
    fPlanActif := Q.FindField ('I_PLANACTIF').AsInteger;
    fMontantHT := Q.FindField ('I_MONTANTHT').AsFloat+
                  Q.FindField ('I_TVARECUPERABLE').AsFloat-Q.FindField ('I_TVARECUPEREE').AsFloat;
    fTVARecuperable := Q.FindField ('I_TVARECUPERABLE').AsFloat;
    fTVARecuperee := Q.FindField ('I_TVARECUPEREE').AsFloat;
    fMttHT:=Q.FindField ('I_MONTANTHT').AsFloat ;
    fBaseTP := Q.FindField ('I_BASETAXEPRO').AsFloat;
    fBaseEco := Q.FindField ('I_BASEECO').AsFloat;
    fBaseFisc := Q.FindField ('I_BASEFISC').AsFloat;
    fResiduel := Q.FindField ('I_RESIDUEL').AsFloat;
    fCompteImmo := Q.FindField ('I_COMPTEIMMO').AsString;
    fCompteCharge := Q.FindField ('I_COMPTELIE').AsString;
    fDureeEco := Q.FindField ('I_DUREEECO').AsInteger;
    fMethodeEco := Q.FindField ('I_METHODEECO').AsString;
    fTauxEco := Q.FindField ('I_TAUXECO').AsFloat;
    fLibelle := Q.FindField ('I_LIBELLE').AsString;
    fQuantite := Q.FindField ('I_QUANTITE').AsFloat;
    fNature := Q.FindField ('I_NATUREIMMO').AsString;
    bFiscal := (Q.FindField ('I_METHODEFISC').AsString <> '');
    fSaveDate:=Q.FindField('I_DATEMODIF').AsDateTime ;
    fEtat := Q.FindField('I_ETAT').AsString;
    fDateLevOpt := Q.FindField('I_DATEFINCB').AsDateTime;
    // BTY 03/06 FQ 17446
    fJournala :=  Q.FindField('I_JOURNALA').AsString;
    // MBO 15/06/2006 - FQ 18395
    fDateDebEco := Q.FindField('I_DATEDEBECO').AsDateTime;
    // BTY 10/06 Prime d'équipement
    fSBVPRI := Q.FindField ('I_SBVPRI').AsFloat;
    // BTY 11/06 Subvention d'investissement
    fSBVMT := Q.FindField ('I_SBVMT').AsFloat;
    // BTY 04/07 nelle gestion fiscale
    fbDPI := (Q.FindField ('I_DPI').AsString = 'X');
    fbFuturVNF := (Q.FindField ('I_FUTURVNFISC').AsString = '***');
    fbGestionfiscale := (Q.FindField ('I_NONDED').AsString = 'X');
    // mbo 12/06/2007 fRemplace := Q.FindField ('I_REMPLACE').AsString;
    fRemplace := Q.FindField ('I_STRING1').AsString;
    fRepriseDR := Q.FindField ('I_REPRISEDR').AsFloat;
    fRepriseFEC := Q.FindField ('I_REPRISEFEC').AsFloat;
    fDateDebFis := Q.FindField('I_DATEDEBFIS').AsDateTime;

    if (fNature='CB') or (fNature='LOC') then
    begin
      fContrat := TImContrat.Create;
      fContrat.Charge (Q);
      fContrat.ChargeTableEcheance;

      // fq 17339 - mbo - 20.01.2006 - stockage de la date fin si levée d'option
      if (fDateFinContrat <> fDateLevOpt) and (DateToStr(fDateLevOpt) <> stDate1900) then //'01/01/1900') then BTY 07/07
          fDateFinCOntrat := fDateLevOpt
      else
          fDateFinContrat := fContrat.GetDateFinContrat;

      fContrat.Free;
    end;
    laI_IMMO.Caption := fCode;
    laI_LIBELLE.Caption := fLibelle;
    laI_DATEPIECEA.Caption := DateToStr(fDateAchat);
    //laI_MONTANTHT.Caption := StrfMontant(fMontantHT,20, V_PGI.OkDecV , '', false);
    if fBaseFisc<>0 then laI_MONTANTHT.Caption := StrfMontant(fBaseFisc,20, V_PGI.OkDecV , '', false)
                    else laI_MONTANTHT.Caption := StrfMontant(fBaseEco,20, V_PGI.OkDecV , '', false);
    laI_COMPTEIMMO.Caption := fCompteImmo;
    laI_QUANTITE.Caption := StrfMontant(fQuantite,20, V_PGI.OkDecV , '', false);
  end;
  Ferme (Q);
  {$IFDEF SERIE1}
    DATEOPE.Text:=datetostr(V_PGI.DateEntree) ; //YCP 12/06/2006
  {$ELSE}
    DATEOPE.text := '';
  {$ENDIF SERIE1}
  fOrdre := TrouveNumeroOrdreLogSuivant(fCode);
  fOrdreS := TrouveNumeroOrdreSerieLogSuivant;
end;

function TFAncOpe.ControleZones : boolean;
var Q : TQuery;
begin
  // La date est-elle valide ?
  if (not IsValidDate (DATEOPE.Text)) then
  begin
    HM.Execute(4,Caption,'');
    FocusControl (DATEOPE);Result := false; exit;
  end;
  // La date est-elle postérieure à la date d'achat ?
  if (StrToDate(DATEOPE.Text) < fDateAchat) then
  begin
    HM.Execute (0,Caption,'');
    FocusControl (DATEOPE); Result := false;exit;
  end;
  // La date est-elle un date de l'exercice en cours ?
  if (StrToDate(DATEOPE.Text) < VHImmo^.Encours.Deb) or ( StrToDate(DATEOPE.Text) > VHImmo^.Encours.Fin) then
  begin
    HM.Execute (1,Caption,'');
    FocusControl (DATEOPE); Result := false;exit;
  end;
  // La date est-elle postérieure à la date de dernière opération ?
  Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+fCode+'" ORDER BY IL_DATEOP DESC', True);
  if not Q.Eof then
    if Q.FindField ('IL_DATEOP').AsDateTime > StrToDate(DATEOPE.Text) then
    begin
      HM.Execute (7,Caption,'');
      FocusControl (DATEOPE);Result := false; exit;
    end;
  Ferme (Q);
  Result := True;
end;

procedure TFAncOpe.bValiderClick(Sender: TObject);
begin
  ModalResult := mrNone;
  NextPrevControl(Self);
  if ControleZones then EnregistreOperation (fProcEnreg);
end;

procedure TFAncOpe.bFermerClick(Sender: TObject);
var mr : integer;
begin
  mr := HM.Execute (3,Caption,'');
  if mr = mrYes then
  begin
    if ControleZones then EnregistreOperation (fProcEnreg)
    else ModalResult := mrNone;
  end
  else if mr = mrNo then ModalResult := mrNo
  else ModalResult := mrNone;
end;

procedure TFAncOpe.EnregistreOperation (PP : TProc);
Var ii : TIoErr ;
begin
  ModalResult := mrNone;
  ii:=Transactions(PP, 1) ;
  if ii=oeSaisie then MessageAlerte(HM.Mess[6])
  else if ii<>oeOK then MessageAlerte(HM.Mess[5])
  else begin
  ModalResult := mrYes;
  VHImmo^.ChargeOBImmo := True;
  ImMarquerPublifi(True);   // CA le 06/10/2000
  end;
end;

procedure TFAncOpe.DATEOPEKeyPress(Sender: TObject; var Key: Char);
begin
 ParamDate (Self, Sender, Key);
end;

procedure TFAncOpe.DATEOPEExit(Sender: TObject);
var Q : TQuery;
begin
  if (StrToDate(DATEOPE.Text) < fDateAchat) then
  begin
    HM.Execute (0,Caption,'');
    FocusControl (DATEOPE); exit;
  end;
  if (StrToDate(DATEOPE.Text) < VHImmo^.Encours.Deb) or ( StrToDate(DATEOPE.Text) > VHImmo^.Encours.Fin) then
  begin
    HM.Execute (1,Caption,'');
    FocusControl (DATEOPE); exit;
  end;
  // La date est-elle postérieure à la date de dernière opération ?
  Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+fCode+'" ORDER BY IL_DATEOP DESC', True);
  if not Q.Eof then
    if Q.FindField ('IL_DATEOP').AsDateTime > StrToDate(DATEOPE.Text) then
    begin
      HM.Execute (7,Caption,'');
      FocusControl (DATEOPE); //EPZ 06/11/00 exit;
    end;
  Ferme (Q);
end;

procedure TFAncOpe.ToolbarButton976Click(Sender: TObject);
begin
CallHelpTopic(self);
end;

procedure TFAncOpe.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_F10 then bValiderClick(nil);
end;

end.
