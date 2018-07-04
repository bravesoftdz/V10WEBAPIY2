{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 01/01/2004
Modifié le ... : 01/01/2004
Description .. : Option Modification des bases d'une immo
Suite......... : FQ 16257 BTY 08/05 Location longue durée non amort. non concernée par la modif bases
Suite......... : CRC 2002-10 BTY 09/05 Ajout d'une check-box protégée par IFDEF car commercialisé plus tard
Suite......... : CRC 2002-10 BTY 10/05 Mise en oeuvre plan fiscal CRC 2002-10
Suite......... : MBO - 19/10/2005 - FQ 16756 - ajout message 10 dans la dfm + test
                 la base éco ne doit pas être inférieure à la base d'amort fiscalement déductible
Suite......... : MBO - 19/10/2005 - FQ 16910 - la base éco ne doit pas être inférieure aux amort antérieurs
Suite......... : BTY - 12/05 - FQ 17171 - Message d'erreur si pb durant la transaction d'enregistrement
Suite......... : BTY - 03/06 - FQ 17446 - Recalculer ou non les antérieurs ECO sur gestion plan fiscal CRC 2002-10
Suite......... : BTY - 03/06 - FQ 17711 - Montant HT et non pas Montant HT + TVA R - TVA r
Suite......... : MBO - 05/06 - FQ 17569 - modif pour dates debut d'amortissement
Suite......... : BTY - 06/06 - FQ 18393 - En série, reprendre la date de l'opération
Suite......... : BTY - 06/06 - FQ 18394  - En série, reprendre le no de serie pour qu'en annulation on puisse annuler les autres opérations de la série
Suite......... : BTY - 07/06 - Tenir compte de la DPI dans le contrôle de la base éco
Suite......... : MBO - 09/06 - Tenir compte de la prime dans le controle base éco ou fisc suivant le cas
Suite......... : BTY - 10/06 - Plan fiscal lié à la présence d'une prime d'équipement
Suite......... : BTY - 10/06 - FQ 18991 Oter un point dans le message 10
Suite......... : BTY - 10/06 - Modif messages 14 et 15
Suite......... : BTY - 10/06 - Précocher Plan fiscal si présence prime d'équipement
Suite......... : BTY - 10/06 - Tenir compte de la subvention dans les contrôles
Suite......... : BTY - 11/06 - Cocher et bloquer Plan fiscal si présence prime ou subvention
Suite......... : BTY - 04/07 - Nelle gestion fiscale, supprimer la notion de plan fiscal CRC200210
Suite......... : BTY - 05/07 - FQ 20256 En modification de base et annulation, mettre le type de dérogatoire
Suite......... : MBO - 12/06/07 - (rev9) pas de modif plan fiscal et base fiscale si immo remplaçante
Suite......... : BTY - 07/07 - Delete d'éléments de typebasefiscale inopérant si libellé en dur => rechdom
Suite......... : BTY - 09/07 FQ 21502 Type base fiscale présente Rempl. Composant => remplacer items.delete par la propriété Plus qui filtre
Suite......... : Va avec le paramétrage &#@ de la tablette TITYPEBASEFISC
Suite......... : MBO - 29/10/2007 - fq 21754 - appel planInfo.calcul avec nveaux paramètres 
Mots clefs ... : IMMO
*****************************************************************}

unit ImModBas;


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
  ImAncOpe,
  StdCtrls,
  Hctrls,
  Mask,
  HSysMenu,
  hmsgbox,
  ComCtrls,
  HRichEdt,
  HRichOLE,
  HTB97,
  ExtCtrls,
  HPanel,
  OpEnCour,
  HEnt1,
  db,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  {$IFDEF VER150}
   Variants,
  {$ENDIF}
  ImPlan,
  ImPlanInfo,
  ImPlanMeth,
  utob,
  ImEnt;

type
  TFModBas = class(TFAncOpe)
    HLabel5: THLabel;
    MONTANTHT: THNumEdit;
    HLabel16: THLabel;
    BASEECO: THNumEdit;
    tBASEFISC: THLabel;
    BASEFISC: THNumEdit;
    tlaI_BASEECO: THLabel;
    laI_BASEECO: THLabel;
    tlaI_BASEFISC: THLabel;
    laI_BASEFISC: THLabel;
    CBPFCRC200210: TCheckBox;
    TYPEBASEFISCALE: THValComboBox;
    BaseFiscaleEgale: THLabel;
    bDureeRest: TCheckBox;
    bDureeRestF: TCheckBox;
    bGestionFiscale: TCheckBox;   // BTY 03/06 FQ 17446
    procedure OnModifZone(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CBPFCRC200210Click(Sender: TObject);
    procedure TYPEBASEFISCALEChange(Sender: TObject);
  private
    { Déclarations privées }
    //fPlanEnCours : TPlanAmort;
    fRepriseEco, fRepriseFisc : double;
    fCalculVNC : boolean;   // BTY 03/06 FQ 17446
    fOrdreSerie : integer;  // BTY FQ 18394
    fFiscalCoche : boolean;
    procedure EnregistreModifBases;
    // BTY 11/05 Transaction pour l'enregistrement dans les tables
    procedure EnregistreTables;
    //procedure ExecuteChangementPlan ( MontantReprise : double);
    procedure DetermineGestionFiscale; // 04/07 Nelle gestion fiscale
  public
    { Déclarations publiques }
    procedure InitZones;override;
    function  ControleZones : boolean;override;
  end;

//function ExecuteModificationBases(Code : string) : TModalResult;
function ExecuteModificationBases(var Code : string) : TModalResult;
procedure AnnuleModifBases(LogModifBases : TLogModifBases; TabMessErreur : THMsgBox);

implementation

uses Outils,ChanPlan,ImSortie,ImOuPlan;

{$R *.DFM}

// FQ 18393
//function ExecuteModificationBases(Code : string) : TModalResult;
function ExecuteModificationBases(var Code : string) : TModalResult;
var FModBas: TFModBas;
begin
  FModBas:=TFModBas.Create(Application) ;
  // FQ 18393
  //FModBas.fCode:=Code ;
  FModBas.fCode:= ReadTokenSt(Code);
  FModBas.fDateSerie := iDate1900;
  if Code <> '' then
     FModBas.fDateSerie:= StrToDate(ReadTokenSt(Code));
  //
  // FQ 18394 2e paramètre : le n° d'ordre en série
  FModBas.fOrdreSerie:= TrouveNumeroOrdreSerieLogSuivant;
  if Code <> '' then
     FModBas.fOrdreSerie:= StrToInt(ReadTokenSt(Code));

  FModBas.fProcEnreg := FModBas.EnregistreModifBases;
  try
    FModBas.ShowModal ;
  finally
    result := FModBas.ModalResult;
    // FQ 18393
    if result = mrYes then
        begin
        Code := FModBas.DATEOPE.Text;
        // FQ 18394
        Code := Code + ';' + IntToStr(FModBas.fOrdreSerie);
       end;
    FModBas.Free ;
  end ;
end ;

procedure TFModBas.InitZones;
begin
  inherited;
  laI_BASEECO.Caption := StrfMontant(fBaseEco,20, V_PGI.OkDecV , '', false);
  tlaI_BASEFISC.Visible := bFiscal;
  laI_BASEFISC.Visible := bFiscal;
  BASEFISC.Visible := bFiscal;
  tBASEFISC.Visible := bFiscal;
  if bFiscal then
    laI_BASEFISC.Caption := StrfMontant(fBaseFisc,20, V_PGI.OkDecV , '', false);
  //pgr 08/2005 FQ 16292 / Gestion décimales
  MONTANTHT.Masks.PositiveMask := StrFMask(V_PGI.OkDecV,'',True);
  BASEECO.Masks.PositiveMask := StrFMask(V_PGI.OkDecV,'',True);
  BASEFISC.Masks.PositiveMask := StrFMask(V_PGI.OkDecV,'',True);

  MONTANTHT.Value := fMttHT; // fMontantHT; // BTY 03/06 FQ 17711 Montant ht sans tva
  MONTANTHT.Enabled := (fEtat = 'OUV');   // FQ 13509
  BASEECO.Value := fBaseEco;
  BASEFISC.Value := fBaseFisc;

  // MVG 09/05
  TypeBaseFiscale.Visible:=FALSE;
  BaseFiscaleEgale.Visible:=FALSE;
  CBPFCRC200210.Visible := false;
  // MVG 09/05

  // BTY 09/05 CheckbB visible si plan fiscal ou méthode éco linéaire ou dégressif
  {*CBPFCRC200210.Visible := ((bFiscal) or (fMethodeEco='LIN') or (fMethodeEco='DEG'));
  CBPFCRC200210.checked := bFiscal;
  CBPFCRC200210.enabled := not (bFiscal);
  if (fBaseEco=fBaseFisc) then
  begin  TYPEBASEFISCALE.Value:='ECO';  end
  else
  begin  TYPEBASEFISCALE.Value:='THE';  end;*}
  // Montant HT non modifiable
  MONTANTHT.Enabled := False;
  // CheckB visible si immo autre que crédit-bail ET soit plan fiscal soit méthode éco linéaire ou dégressif
  // soit présence prime soit présence subvention  BTY 11/06
  CBPFCRC200210.Visible := ( ((fNature<>'CB') and (bFiscal or (fMethodeEco='LIN') or (fMethodeEco='DEG')))
                          or (fSBVPRI <> 0) or (fSBVMT <> 0) );

  // FQ 21502 Items.Delete met du désordre dans les valeurs
  if (fRemplace='') then
     // Remonté ICI AVANT CBPFCRC200210.Checked qui alimente itemindex
     // Car Plus remet itemindex à -1
     TypeBaseFiscale.Plus :=  ' AND CO_CODE<>"TR2"';


  if CBPFCRC200210.Visible then
     begin
     // CheckB bloquée et cochée si plan fiscal ou prime ou subvention  BTY 11/06
     CBPFCRC200210.Checked := (bFiscal or (fSBVPRI <> 0) or (fSBVMT <> 0));
     CBPFCRC200210.Enabled := not CBPFCRC200210.Checked;
     // ComboB visible si plan fiscal
     TYPEBASEFISCALE.Visible := CBPFCRC200210.Checked; // (bFiscal); prime
     BaseFiscaleEgale.Visible := CBPFCRC200210.Checked; // (bFiscal); prime
     end;

  // 04/07 Composant remplaçant => affichage RC2 dans le type de base, sinon le retirer
  if (fRemplace='') then
     //BTY - 07/07 - Delete d'éléments est inopérant si libellé en dur => rechdom
     // FQ 21502 voir plus haut
     //TypeBaseFiscale.Items.Delete(TypeBaseFiscale.Items.IndexOf(RechDom('TITYPEBASEFISC', 'TR2', FALSE)));//'Rempl. composant 2ème catégorie'));
  else
  begin
     TypeBaseFiscale.ItemIndex := TypeBaseFiscale.Items.IndexOf(RechDom('TITYPEBASEFISC', 'TR2', FALSE)) ;
     // ajout mbo 12.06.07 - rev9 - ajout si immo remplaçante pas de modif du plan fiscal
     TYPEBASEFISCALE.Enabled := false;
  end;
  // BTY 03/06 FQ 17446
  // Calcul plan futur avec la VNC si déjà pratiqué ou si plan fiscal coché
  // ou si méthode LIN ou VAR
  fCalculVnc := (fJournala = '***');
  bDureeRest.Visible := ( (fJournala = '***') or
                          ((fMethodeEco='LIN') or (fMethodeEco='VAR')) or
                          ((CBPFCRC200210.Visible) and (CBPFCRC200210.Checked)
                            and (fMethodeEco<>'DEG')) );
  bDureeRest.Checked := ((bDureeRest.Visible) and (fJournala = '***') );
  //
  // BaseFisc toujours bloquée
  BASEFISC.Visible := CBPFCRC200210.Checked; //(bFiscal); prime
  tBaseFisc.Visible:= CBPFCRC200210.Checked; //(bFiscal); prime
  BASEFISC.Enabled := False;
  //{$ENDIF}
  // FQ 18393
  if fDateSerie <> iDate1900 then
     DATEOPE.Text := DateToStr(fDateSerie);
  // 04/07
  bDureeRestF.Visible := CBPFCRC200210.Checked;

  bDureeRestF.Enabled := bDureeRestF.Visible;
  bDureeRestF.Checked := (bDureeRestF.Enabled and fbFuturVNF);

  // ajout mbo - rev9 - pas de modif du plan fiscal si immo remplaçante
  if (fRemplace<>'') then
     bDureeRestF.Enabled := false;

  bGestionFiscale.Visible := CBPFCRC200210.Checked;
  bGestionFiscale.Checked := (CBPFCRC200210.Checked and fbGestionfiscale);

  // mbo - rev9 - ajout du test sur fremplace car pas de modif du plan fiscal si immo remplaçante
  bGestionFiscale.Enabled := (bGestionFiscale.Visible and bFiscal and (fremplace = ''));
  DetermineGestionFiscale;
  fFiscalCoche := CBPFCRC200210.Checked;
end;

function TFModBas.ControleZones : boolean;
var
    PlanInfo : TPlanInfo;
    QPlan : TQuery;
    TPlan : TPlanAmort;
    CumulEco, CumulFisc : double;
begin
  // BTY 08/05 FQ 16257 Location longue durée non amortissable non concernée par la modif bases
  // => Eviter le ControleZones de FAncOpe qui vérifie les dates
  // => et pouvoir sortir sans valider sans message

  // MVG 08/05 FQ 16527 Annulation de la correction,
  // il faut en fait ne pas pouvoir prendre l'option de Modification des bases
  //  if fNature = 'LOC' then
  //  begin
  //  ModalResult := mrNone;Result := false; exit;
  //  end;
  //
  result := inherited ControleZones;
  if result = false then exit;

  if (fBaseEco = BASEECO.Value)
      and (fFiscalCoche = CBPFCRC200210.Checked)
      and (fBaseFisc = BASEFISC.Value)
      and (fCalculVnc = bDureeRest.Checked)
      and (fbFuturVNF = bDureeRestF.Checked)
      and (fDateDebEco = fDateDebFis)
      and (not bGestionFiscale.Checked) then
  begin
    HM.Execute (18,Caption,'');
    ModalResult := mrNone;Result := false; exit;

  // BTY 03/06 FQ 17711 Montant ht sans tva
  //if (MONTANTHT.Value = fMontantHT) and (BASEECO.Value = fBaseEco)
  end else if (MONTANTHT.Value = fMttHT) and (BASEECO.Value = fBaseEco)
       and (fFiscalCoche = CBPFCRC200210.Checked)
       and (BASEFISC.Value = fBaseFisc)
       and (fCalculVnc = bDureeRest.Checked)
       and (fbFuturVNF = bDureeRestF.Checked) then
// BTY 03/06 FQ 17446   and (BASEFISC.Value = fBaseFisc) then
  begin
    HM.Execute (5,Caption,'');
    ModalResult := mrNone;Result := false; exit;


  end else if (BASEECO.Value > (MONTANTHT.Value + fTVARecuperable - fTvaRecuperee)) then
  begin
    HM.Execute (6,Caption,'');
    ModalResult := mrNone;Result := false; exit;
  end else
  begin
    { Si l'immo est totalement amortie, on interdit la modification des bases }
    PlanInfo:= TPlanInfo.Create (fCode);
    PlanInfo.ChargeImmo(fCode);
    PlanInfo.Calcul(StrToDate(DATEOPE.Text),True, false, '');   // fq 21754
    if PlanInfo.VNCEco = 0 then
    begin
        HM.Execute (8,Caption,'');
        ModalResult := mrNone;
        Result := false;
        PlanInfo.Free;
        exit;
    end else PlanInfo.Free;
  end;

  // modif mbo FQ 16756 + fq 16910
  // BTY 10/05 Contrôle supplémentaire en CRC 2002-10
  //        if (bFiscal) then
  //        begin
  TPlan:=TPlanAmort.Create(true) ;
  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+ fCode +'"', FALSE) ;
  TPlan.Charge(QPlan);
  TPlan.Recupere(fCode,QPlan.FindField('I_PLANACTIF').AsString);
  TPlan.GetCumulsDotExercice(VHImmo^.Encours.Deb,CumulEco, CumulFisc,true,true,true);

  // Bloquer si on a coché Gestion plan fiscal CRC 2002-10 alorsque l'immo a été dépréciée
  if ((CBPFCRC200210.Checked) and (CBPFCRC200210.Enabled) and
      ((TPlan.GestDeprec) or (TPlan.DureeRest))) then
  begin
      HM.Execute (12,Caption,'');
      ModalResult := mrNone; Result := false;
      TPlan.Free ; Ferme(QPlan);
      exit;
  end;
  // ajout mbo fq 16756
  // BTY 07/06 DPI
  //if (BASEECO.Value < TPlan.ValReintegration) then
  if (BASEECO.Value - TPlan.ValReintegration - TPlan.DPIAffectee) < 0 then
  begin
      HM.Execute (10,Caption,'');
      ModalResult := mrNone; Result := false;
      TPlan.Free ; Ferme(QPlan);
      exit;
  // ajout mbo FQ 16910
  end else if (CumulEco > BaseEco.Value) then
      begin
      HM.Execute (11,Caption,'');
      ModalResult := mrNone; Result := false;
      TPlan.Free ; Ferme(QPlan);
      exit;
  end;
  if (bFiscal) and (CumulFisc > BaseFisc.Value) then
    begin
    HM.Execute (9,Caption,'');
    ModalResult := mrNone; Result := false;
    TPlan.Free ; Ferme(QPlan);
    exit;
    end;

  // ajout mbo - 18.09.06 pour controle du montant de la prime
  // qui ne doit pas être supérieure à la base prise en compte
  if (bFiscal) and (Tplan.MNTPrime > BaseFisc.Value) then
  begin
     HM.Execute (14,Caption,'');
     ModalResult := mrNone; Result := false;
     TPlan.Free ; Ferme(QPlan);
     exit;
  end else
  begin
     if (Tplan.MNTPrime > BaseEco.Value) then
     begin
        HM.Execute (15,Caption,'');
        ModalResult := mrNone; Result := false;
        TPlan.Free ; Ferme(QPlan);
        exit;
     end;
  end;

  // BTY Prise en compte Subvention d'investissement
  if (bFiscal) and (Tplan.MNTSbv > BaseFisc.Value) then
  begin
     HM.Execute (16,Caption,'');
     ModalResult := mrNone; Result := false;
     TPlan.Free ; Ferme(QPlan);
     exit;
  end else
  begin
     if (Tplan.MNTSbv > BaseEco.Value) then
     begin
        HM.Execute (17,Caption,'');
        ModalResult := mrNone; Result := false;
        TPlan.Free ; Ferme(QPlan);
        exit;
     end;
  end;

  // BTY 03/06 FQ 17446
  if not bDureeRest.Visible then bDureeRest.Checked := False;
  //
  TPlan.Free; Ferme(QPlan);

end;


procedure TFModBas.EnregistreModifBases;
begin
  // Transaction pour ne pas enregister une table sans l'autre
  BEGINTRANS ;
  try
    EnregistreTables;
    COMMITTRANS ;
  except
    HM.Execute (13,Caption,'');
    ROLLBACK ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BTY
Créé le ...... : 08/11/2005
Modifié le ... :   /  /
Description .. : Isoler les traitements dans les 3 tables IMMO, IMMOAMOR
Suite ........ : et IMMOLOG dans une même transaction pour garder la
Suite ........ : cohérence en cas de clash
Mots clefs ... :
*****************************************************************}
procedure TFModBas.EnregistreTables;
var QImmo : TQuery;
    Plan, Plan2 : TPlanAmort;
    PlanActif : integer;
    CumulEco,CumulFisc : double;
    MethodeFiscale : string;
    DureeFiscale : integer;
    TLog : TOB;
begin
  PlanActif := 0;
  DureeFiscale := 0;

  Plan:=TPlanAmort.create(true) ;// := CreePlan (True);

  // Plan fiscal coché
  if CBPFCRC200210.checked  then
    begin

    try
      QImmo := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+fCode+'"',False);
      if not QImmo.Eof then
      begin
         fRepriseEco := QImmo.FindField('I_REPRISEECO').AsFloat;
         fRepriseFisc := QImmo.FindField('I_REPRISEFISCAL').AsFloat;
         Plan.Charge (QImmo);
         Plan.Recupere(fCode,QImmo.FindField('I_PLANACTIF').AsString);
         Plan.GetCumulsDotExercice(VHImmo^.Encours.Deb, CumulEco,CumulFisc,True,True,true);
         // Pour recalculer les antérieurs ECO
         Plan2:=TPlanAmort.create(true) ;
//         Plan2.copie(Plan);

         Plan.free ;
         Plan:=TPlanAmort.Create(true) ;// := CreePlan ( True);;
         QImmo.Edit;
         QImmo.FindField ('I_MONTANTHT').AsFloat := MONTANTHT.Value;
         QImmo.FindField ('I_BASEECO').AsFloat := BASEECO.Value;
         QImmo.FindField ('I_BASEFISC').AsFloat := BASEFISC.Value;
         MethodeFiscale := QImmo.FindField ('I_METHODEFISC').AsString;
         DureeFiscale := QImmo.FindField ('I_DUREEFISC').AsInteger;
         if (not (bFiscal)) then
            begin
            QImmo.FindField ('I_METHODEFISC').AsString := QImmo.FindField ('I_METHODEECO').AsString;
            QImmo.FindField ('I_DUREEFISC').AsInteger := QImmo.FindField ('I_DUREEECO').AsInteger;
            QImmo.FindField ('I_TAUXFISC').AsFloat := QImmo.FindField ('I_TAUXECO').AsFloat;
            QImmo.FindField('I_REPRISEFISCAL').AsFloat := CumulEco;
            end;
         // BTY 03/06 FQ 17446 Pas de recalcul des antérieurs si Calcul plan futur sur la VNC
         if bDureeRest.Checked then
            begin
            QImmo.FindField('I_REPRISEECO').AsFloat := CumulEco;
            QImmo.FindField('I_JOURNALA').AsString := '***';
            end
            else
            begin
         // BTY 03/06 FQ 17446
         QImmo.FindField('I_JOURNALA').AsString := '';
         //
         // La Reprise ECO doit être recalculée en tenant compte de la nelle base ECO
         // (pas la Reprise FISC)
         QImmo.FindField('I_REPRISEECO').AsFloat := 0 ;
         Plan2.Charge (QImmo);
         Plan2.AmortEco.Reprise := 0;
         // BTY 03/06  FQ 17500
         Plan2.AmortEco.RepriseDep := 0;
         Plan2.AmortFisc.RepriseDep := 0;
         //
         Plan2.AmortEco.Base := BASEECO.Value;
         Plan2.CalculDateFinAmortReprise(Plan2.AmortEco);
         // BTY 03/06  Antérieurs/Amortissements faux en dégressif si tableauDOT non réinitalisé
//         Plan2.ResetTableauDot(Plan2.AmortEco, Plan2.GetDateDebutAmort (Plan2.AmortEco) );
         Plan2.ResetTableauDot(Plan2.AmortEco, Plan2.DateDebEco);

         //
         Plan2.CalculReprise(Plan2.AmortEco);
         QImmo.FindField('I_REPRISEECO').AsFloat :=Plan2.AmortEco.Reprise;
            end;
         //
         Plan2.free ;
         // Eléments déjà amortis
         if (bFiscal) then QImmo.FindField('I_REPRISEFISCAL').AsFloat := CumulFisc;
         // Les antérieurs ECO ne doivent pas dépasser les antérieurs FISCAUX
         {if (QImmo.FindField('I_REPRISEECO').AsFloat > QImmo.FindField('I_REPRISEFISCAL').AsFloat) then
             QImmo.FindField('I_REPRISEECO').AsFloat := QImmo.FindField('I_REPRISEFISCAL').AsFloat; }

         // 04/07 Nelle gestion fiscale : plan futur sur la VNF
         if bDureeRestF.Checked then
            QImmo.FindField('I_FUTURVNFISC').AsString := '***'
         else
            QImmo.FindField('I_FUTURVNFISC').AsString := '';
         // 04/07 Nelle gestion fiscale : Gestion fiscale
         if bGestionFiscale.Checked then
            QImmo.FindField('I_NONDED').AsString := 'X'
         else
            QImmo.FindField('I_NONDED').AsString := '-';

         QImmo.FindField('I_TYPEDEROGLIA').AsString := TypeDerogatoire( nil, QImmo); // FQ 20256

         CocheChampOperation (QImmo,fCode,'I_OPEMODIFBASES');
         Plan.Charge (QImmo);

         // Antérieurs ECO recalculés => recalculer antérieurs dérogatoires
         if (not bDureeRest.Checked) then
         begin
           Plan.CalculReprise_Derog_Reint;
           QImmo.FindField('I_REPRISEDR').AsFloat := Plan.AmortDerog.Reprise;
           QImmo.FindField('I_REPRISEFEC').AsFloat := Plan.AmortReint.Reprise;
         end;

         Plan.Calcul (QImmo,StrToDate(DATEOPE.Text));
         Plan.TypeOpe := 'MB2'; // BTY 10/05 Enregistrer le code opération ds IMMOAMOR
         Plan.Sauve;
         PlanActif := Plan.NumSeq;
         QImmo.FindField('I_PLANACTIF').AsInteger := PlanActif;
         QImmo.Post;
      end;
      Ferme (QImmo);
    finally
      Plan.free ; //Detruit;
    end ;

    {
    QueryW:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+W_W+'"', FALSE) ;
    QueryW.Insert ;
    QueryW.FindField('IL_IMMO').AsString:=fCode;
    QueryW.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('MBA'); //15/01/99 EPZ
    QueryW.FindField('IL_DATEOP').AsDateTime:=StrToDate(DATEOPE.Text) ;
    QueryW.FindField('IL_TYPEOP').AsString:='MB2';
    QueryW.FindField('IL_LIBELLE').AsString := RechDom('TIOPEAMOR', 'MB2', FALSE)+' '+DATEOPE.text;
    QueryW.FindField('IL_METHODEFISC').AsString:= MethodeFiscale;
    QueryW.FindField('IL_DUREEFISC').AsInteger:= DureeFiscale;
    QueryW.FindField('IL_REPRISEECO').AsFloat:= fRepriseEco;
    QueryW.FindField('IL_REPRISEFISC').AsFloat:= fRepriseFisc;
    // BTY 12/10/05 Compatibilité CWAS
    //TBlobField(QueryW.FindField('IL_BLOCNOTE')).Assign(IL_BLOCNOTE.LinesRTF);
    QueryW.FindField('IL_BLOCNOTE').AsString :=  RichToString (IL_Blocnote);
    QueryW.FindField('IL_PLANACTIFAV').AsFloat:=PlanActif-1;
    QueryW.FindField('IL_PLANACTIFAP').AsFloat:=PlanActif;
    QueryW.FindField('IL_ORDRE').AsInteger:=fOrdre;
    QueryW.FindField('IL_ORDRESERIE').AsInteger:=fOrdreS;
    QueryW.FindField('IL_MONTANTAVMB').AsFloat:=fMontantHT;
    QueryW.FindField('IL_BASEECOAVMB').AsFloat:=fBaseEco;
    QueryW.FindField('IL_BASEFISCAVMB').AsFloat:=fBaseFisc;
    //QueryW.FindField('IL_LIBELLE').AsString := RechDom('TIOPEAMOR', 'MBA', FALSE)+' '+DATEOPE.text;
    QueryW.FindField('IL_CUMANTCESECO').AsFloat:=fRepriseEco;
    QueryW.FindField('IL_CUMANTCESFIS').AsFloat:=fRepriseFisc;
    QueryW.Post;
    Ferme(QueryW) ;
    }

    // TOB plus avantageuse que Query car évite de compléter les champs vides
    TLog := TOB.Create ('IMMOLOG',nil,-1);
    try
      TLog.PutValue('IL_IMMO',fCode);
      TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('MBA'));
      TLog.PutValue('IL_DATEOP', StrToDate(DATEOPE.Text));
      TLog.PutValue('IL_TYPEOP','MB2');
      TLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', 'MB2', FALSE)+' '+DATEOPE.text);
      TLog.PutValue('IL_METHODEFISC', MethodeFiscale);
      TLog.PutValue('IL_DUREEFISC', DureeFiscale);
      TLog.PutValue('IL_REPRISEECO', fRepriseEco);
      TLog.PutValue('IL_REPRISEFISC', fRepriseFisc);
      TLog.PutValue('IL_BLOCNOTE', RichToString (IL_Blocnote));
      TLog.PutValue('IL_PLANACTIFAV', PlanActif-1);
      TLog.PutValue('IL_PLANACTIFAP', PlanActif);
      TLog.PutValue('IL_ORDRE', fOrdre);
      TLog.PutValue('IL_ORDRESERIE', fOrdreSerie); // FQ 18394 fOrdreS);
      TLog.PutValue('IL_MONTANTAVMB', fMttHT); // BTY 03/06 FQ 17711 fMontantHT);
      TLog.PutValue('IL_BASEECOAVMB', fBaseEco);
      TLog.PutValue('IL_BASEFISCAVMB', fBaseFisc);
      TLog.PutValue('IL_CUMANTCESECO', fRepriseEco);
      TLog.PutValue('IL_CUMANTCESFIS', fRepriseFisc);
      // BTY 12/05 Si l'immo avait déjà un plan fiscal, stocker les anciens antérieurs à fin N-1
      if (bFiscal) then
         begin
         TLog.PutValue('IL_MONTANTDOT', CumulEco);
         TLog.PutValue('IL_MONTANTECL', CumulFisc);
         // 04/07 ancien Plan futur avec la VNF
         if fbFuturVNF then
            TLog.PutValue('IL_LIEUGEO', '***')
         else
            TLog.PutValue('IL_LIEUGEO', '');
         // 04/07 anciennne Gestion fiscale
         if fbGestionfiscale then
            TLog.PutValue('IL_ETABLISSEMENT', 'X')
         else
            TLog.PutValue('IL_ETABLISSEMENT', '');
         end;
      // BTY 03/06 FQ 17446  Stocker l'ancien I_JOURNALA
      TLog.PutValue('IL_CODEECLAT', fJournala);
      TLog.PutValue('IL_VERSION', V_PGI.NumVersion);
      // Antérieurs dérogatoires
      TLog.PutValue('IL_REVISIONDOTECO', fRepriseDR);
      TLog.PutValue('IL_REVISIONREPECO', fRepriseFEC);

      TLog.InsertDB(nil);
    finally
      TLog.Free;
    end;

    //{$ENDIF}
    end
    else
    begin

    // BTY 10/05 Modif base habituelle (hors plan fiscal CRC 2002-10)
    try
    QImmo := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+fCode+'"',False);
    if not QImmo.Eof then
      begin
        fRepriseEco := QImmo.FindField('I_REPRISEECO').AsFloat;
        fRepriseFisc := QImmo.FindField('I_REPRISEFISCAL').AsFloat;
        Plan.Charge (QImmo);
        Plan.Recupere(fCode,QImmo.FindField('I_PLANACTIF').AsString);
        Plan.GetCumulsDotExercice(VHImmo^.Encours.Deb, CumulEco,CumulFisc,True,True,true);
        Plan.free ;
        Plan:=TPlanAmort.Create(true) ;// := CreePlan ( True);;
        QImmo.Edit;
        QImmo.FindField ('I_MONTANTHT').AsFloat := MONTANTHT.Value;
        QImmo.FindField ('I_BASEECO').AsFloat := BASEECO.Value;
        QImmo.FindField ('I_BASEFISC').AsFloat := BASEFISC.Value;
        { Mémorisation des éléments déjà amortis }
        QImmo.FindField('I_REPRISEECO').AsFloat := CumulEco;
        QImmo.FindField('I_REPRISEFISCAL').AsFloat := CumulFisc;
        CocheChampOperation (QImmo,fCode,'I_OPEMODIFBASES');
        // BTY 03/06 FQ 17446 Pas de recalcul des antérieurs si Calcul plan futur sur la VNC
        if (bDureeRest.Visible) then
           if bDureeRest.Checked then
              QImmo.FindField ('I_JOURNALA').AsString := '***'
           else
              QImmo.FindField ('I_JOURNALA').AsString := '';
        //
        Plan.Charge (QImmo);
        Plan.Calcul (QImmo,StrToDate(DATEOPE.Text));
        Plan.TypeOpe := 'MBA'; // BTY 10/05 Enregistrer le code opération ds IMMOAMOR
        Plan.Sauve;
        PlanActif := Plan.NumSeq;
        QImmo.FindField('I_PLANACTIF').AsInteger := PlanActif;
        QImmo.Post;
      end;
      Ferme (QImmo);
    finally
      Plan.free ; //Detruit;
    end ;

    {
    QueryW:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+W_W+'"', FALSE) ;
    QueryW.Insert ;
    QueryW.FindField('IL_IMMO').AsString:=fCode;
    QueryW.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('MBA'); //15/01/99 EPZ
    QueryW.FindField('IL_DATEOP').AsDateTime:=StrToDate(DATEOPE.Text) ;
    QueryW.FindField('IL_TYPEOP').AsString:='MBA';
    QueryW.FindField('IL_LIBELLE').AsString := RechDom('TIOPEAMOR', 'MBA', FALSE)+' '+DATEOPE.text;
    // BTY 12/10/05 Compatibilité CWAS
    //TBlobField(QueryW.FindField('IL_BLOCNOTE')).Assign(IL_BLOCNOTE.LinesRTF);
    QueryW.FindField('IL_BLOCNOTE').AsString :=  RichToString (IL_Blocnote);

    QueryW.FindField('IL_PLANACTIFAV').AsFloat:=PlanActif-1;
    QueryW.FindField('IL_PLANACTIFAP').AsFloat:=PlanActif;
    QueryW.FindField('IL_ORDRE').AsInteger:=fOrdre;
    QueryW.FindField('IL_ORDRESERIE').AsInteger:=fOrdreS;
    QueryW.FindField('IL_MONTANTAVMB').AsFloat:=fMontantHT;
    QueryW.FindField('IL_BASEECOAVMB').AsFloat:=fBaseEco;
    QueryW.FindField('IL_BASEFISCAVMB').AsFloat:=fBaseFisc;
    QueryW.FindField('IL_LIBELLE').AsString := RechDom('TIOPEAMOR', 'MBA', FALSE)+' '+DATEOPE.text;
    QueryW.FindField('IL_CUMANTCESECO').AsFloat:=fRepriseEco;
    QueryW.FindField('IL_CUMANTCESFIS').AsFloat:=fRepriseFisc;
    QueryW.Post;
    Ferme(QueryW) ;
    }

    // BTY 11/05 TOB plus avantageuse que Query car évite de compléter les champs vides
    TLog := TOB.Create ('IMMOLOG',nil,-1);
    try
      TLog.PutValue('IL_IMMO',fCode);
      TLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('MBA'));
      TLog.PutValue('IL_DATEOP', StrToDate(DATEOPE.Text));
      TLog.PutValue('IL_TYPEOP','MBA');
      TLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', 'MBA', FALSE)+' '+DATEOPE.text);
      TLog.PutValue('IL_BLOCNOTE', RichToString (IL_Blocnote));
      TLog.PutValue('IL_PLANACTIFAV', PlanActif-1);
      TLog.PutValue('IL_PLANACTIFAP', PlanActif);
      TLog.PutValue('IL_ORDRE', fOrdre);
      TLog.PutValue('IL_ORDRESERIE', fOrdreSerie);  // FQ 18394  fOrdreS);
      TLog.PutValue('IL_MONTANTAVMB', fMttHT); // BTY 03/06 FQ 17711 fMontantHT);
      TLog.PutValue('IL_BASEECOAVMB', fBaseEco);
      TLog.PutValue('IL_BASEFISCAVMB', fBaseFisc);
      TLog.PutValue('IL_CUMANTCESECO', fRepriseEco);
      TLog.PutValue('IL_CUMANTCESFIS', fRepriseFisc);
      // BTY 03/06 FQ 17446  Stocker l'ancien I_JOURNALA
      TLog.PutValue('IL_CODEECLAT', fJournala);
      // 04/07 Nelle gestion fiscale
      if fbFuturVNF then
         TLog.PutValue('IL_LIEUGEO', '***')
      else
         TLog.PutValue('IL_LIEUGEO', '');
      if fbGestionfiscale then
         TLog.PutValue('IL_ETABLISSEMENT', 'X')
      else
         TLog.PutValue('IL_ETABLISSEMENT', '');
      TLog.PutValue('IL_VERSION', V_PGI.NumVersion);

      // Antérieurs dérogatoires
      TLog.PutValue('IL_REVISIONDOTECO', fRepriseDR);
      TLog.PutValue('IL_REVISIONREPECO', fRepriseFEC);

      TLog.InsertDB(nil);
    finally
      TLog.Free;
    end;

    end; // fin if CBPFCRC

  // Ajout CA le 17/03/99 pour gestion reprise le cas échéant
  // modif mbo - fq 16756 - on n'autorise plus base eco < anterieurs
  {if ( CumulEco - Valeur(BASEECO.Text)) > 0 then
  begin
    fPlanEnCours:=TPlanAmort.Create(true) ;//  := CreePlan (True);
    try
      QImmo := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+fCode+'"',True);
      if not QImmo.Eof then
      begin
        fPlanEnCours.Charge (QImmo);
        fPlanEnCours.Recupere(fCode,IntToStr(PlanActif));
        ExecuteChangementPlan ( CumulEco - Valeur(BASEECO.Text));
        PGIInfo ('Le cumul  étant supérieur à la base d''amortissement, une reprise a été réalisée.',Caption);
      end;
      Ferme(QImmo);
    finally
      fPlanEnCours.free ;//Detruit;
    end ;
  end;
  }
  // fin ajout CA
end;

{*
procedure TFModBas.ExecuteChangementPlan ( MontantReprise : double);
var ChangePlan : TChangePlanAmort; TypeExc : string;
begin
  ChangePlan:=TChangePlanAmort.Create (Application);
  ChangePlan.Plan.copie(fPlanEnCours);
  ChangePlan.AncienPlan.copie(fPlanEnCours);
  ChangePlan.CodeImmo:=fCode;
  ChangePlan.DateOpe :=StrToDate(DATEOPE.Text);
  ChangePlan.TypeChangement := 2;
  ChangePlan.Plan.MontantExc:=ChangePlan.Plan.MontantExc-MontantReprise;
  TypeExc := 'RDO';
  ChangePlan.Plan.TypeOpe := 'ELC';
  ChangePlan.Plan.SetTypeExc(TypeExc);
  ChangePlan.TypeExc := TypeExc;
  ChangePlan.TypeDot := TypeExc;
  ChangePlan.MontantDot := MontantReprise;
  TraiteChangementPlan (ChangePlan.Plan);
  ChangePlan.OrdreS :=fOrdreS; // pour forcer le calcul du Numero d'ordre en serie
  ChangePlan.EcritLogChangePlan;
  ChangePlan.free;
end;
*}

procedure TFModBas.OnModifZone(Sender: TObject);
begin
  inherited;
  if THNumEdit(Sender).Name = 'MONTANTHT' then
  begin
    BASEECO.Value := MONTANTHT.Value;
    //if bFiscal then BASEFISC.Value := MONTANTHT.Value;
    // BTY 10/05 CRC2002-10
    if bFiscal then BASEFISC.Value := BASEECO.Value;
    if typebasefiscale.Value = 'ECO' then BASEFISC.Value := BASEECO.Value
    else if typebasefiscale.Value = 'THE' then BASEFISC.Value := MONTANTHT.Value;
    DetermineGestionFiscale;
  end;
  if THNumEdit(Sender).Name = 'BASEECO' then
  begin
    // if bFiscal then BASEFISC.Value := BASEECO.Value;
    // BTY 10/05 CRC2002-10
    if bFiscal then BASEFISC.Value := BASEECO.Value;
    if typebasefiscale.Value = 'ECO' then BASEFISC.Value := BASEECO.Value
    else if typebasefiscale.Value = 'THE' then BASEFISC.Value :=
                                      MONTANTHT.Value + fTVARecuperable - fTvaRecuperee;
                                      // BTY 03/06 FQ 17711 MONTANTHT.Value;
    DetermineGestionFiscale;
  end;
end;

procedure AnnuleModifBases(LogModifBases : TLogModifBases; TabMessErreur : THMsgBox);
var QueryImmo : TQuery;
    PlanActif : string;
    TauxFiscal : double;
    //DateAmort,DatePieceA : TdateTime;  fq 17569
    DatedebFis : TdateTime;
    NumVersion : string;
begin
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogModifBases.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;
  PlanActif:=QueryImmo.FindField('I_PLANACTIF').AsString ;

  // BTY 10/05 Annulation avec plan fiscal CRC2002-10
  // fq 17569 DateAmort := QueryImmo.FindField('I_DATEAMORT').AsDateTime;
  // fq 17569 DatePieceA := QueryImmo.FindField('I_DATEPIECEA').AsDateTime;

  // mbo fq 17569
  DatedebFis :=QueryImmo.FindField('I_DATEDEBFIS').AsDateTime;

  // mbo fq 17569
  { if (not VarIsNull(DateAmort)) and (not VarIsNull(DatePieceA)) and
   (not VarIsNull(LogModifBases.DureeFisc)) then
        TauxFiscal := GetTaux(LogModifBases.MethodeFisc, DateAmort, DatePieceA,
                              LogModifBases.DureeFisc) }

  if (not VarIsNull(DatedebFis)) and
   (not VarIsNull(LogModifBases.DureeFisc)) then
     TauxFiscal := GetTaux(LogModifBases.MethodeFisc, DatedebFis, DatedebFis,
                              LogModifBases.DureeFisc)
  else  TauxFiscal := 0.0;
  QueryImmo.FindField('I_METHODEFISC').AsString := LogModifBases.MethodeFisc;
  QueryImmo.FindField('I_DUREEFISC').AsInteger := LogModifBases.DureeFisc;
  QueryImmo.FindField('I_TAUXFISC').AsFloat := TauxFiscal;
  //{$ENDIF}

  ExecuteSQL('DELETE FROM IMMOAMOR WHERE (IA_IMMO="'+LogModifBases.Lequel+
                     '" AND IA_NUMEROSEQ='+PlanActif+')') ;
  QueryImmo.FindField('I_MONTANTHT').AsFloat := LogModifBases.MontantHT;
  QueryImmo.FindField('I_BASEECO').AsFloat := LogModifBases.BaseEco;
  QueryImmo.FindField('I_BASEFISC').AsFloat := LogModifBases.BaseFisc;
  QueryImmo.FindField('I_PLANACTIF').AsFloat := LogModifBases.PlanActifAv;
  QueryImmo.FindField('I_REPRISEECO').AsFloat := LogModifBases.RepriseEco;
  QueryImmo.FindField('I_REPRISEFISCAL').AsFloat := LogModifBases.RepriseFisc;
  // BTY 03/06 FQ 17446 Restaurer I_JOURNALA
  QueryImmo.FindField('I_JOURNALA').AsString:=LogModifBases.CodeEclat;
  // 04/07 Plan futur avec la VNF si version 8 ou supérieure
  NumVersion := LogModifBases.NumVersion;
  NumVersion := Copy (NumVersion, 1, POS('.',NumVersion)-1);
  // 04/07 Gestion fiscale
  if (StrToInt ( Copy (NumVersion,1,1)) >= 8) then
  begin
     QueryImmo.FindField('I_FUTURVNFISC').AsString:=LogModifBases.CalculVNF;
     QueryImmo.FindField('I_NONDED').AsString:=LogModifBases.GestionFiscale;
  end;
  QueryImmo.FindField('I_TYPEDEROGLIA').AsString := TypeDerogatoire( nil, QueryImmo); // FQ 20256

  // Antérieurs dérogatoires
  QueryImmo.FindField('I_REPRISEDR').AsFloat := LogModifBases.RepriseDR;
  QueryImmo.FindField('I_REPRISEFEC').AsFloat := LogModifBases.RepriseFEC;

  // BTY Annulation modif base MBA ou MB2 (avec CRC 2002-10)
  // MajOpeEnCoursImmo ( QueryImmo,'I_OPEMODIFBASES', 'MBA', '-');
  MajOpeEnCoursImmo ( QueryImmo,'I_OPEMODIFBASES', LogModifBases.TypeOpe, '-');
  QueryImmo.Post;
  Ferme (QueryImmo);
end;

procedure TFModBas.FormCreate(Sender: TObject);
begin
  inherited;
{$IFDEF SERIE1}
HelpContext:=511030 ;
{$ELSE}
HelpContext:=2111800 ;
{$ENDIF}
end;

// MVG 09/05 et BTY 10/05
procedure TFModBas.CBPFCRC200210Click(Sender: TObject);
begin
  inherited;
  // BTY 03/06 FQ 17446
  bDureeRest.Visible := ( (fJournala='***') or
                          ((CBPFCRC200210.Checked) and (fMethodeEco<>'DEG')) or
                          ((fMethodeEco='LIN') or (fMethodeEco='VAR'))  );
  if not bDureeRest.Visible then bDureeRest.Checked := False;

  if CBPFCRC200210.checked then
     begin
     typebasefiscale.visible:=true;
     tbasefisc.visible:=true;
     basefisc.visible:=true;
     BaseFiscaleEgale.Visible:=true;
     bDureeRestF.Visible := True;
     bDureeRestF.Enabled := True;
     bGestionFiscale.Visible := True;

     if bFiscal then
        begin
        if (fBaseEco=fBaseFisc) then
           begin
           typebasefiscale.itemindex:=typebasefiscale.items.indexof(RechDom('TITYPEBASEFISC','ECO',FALSE));
           BaseFisc.Value := BaseEco.Value;
           end
        else
           begin
           typebasefiscale.itemindex:=typebasefiscale.items.indexof(RechDom('TITYPEBASEFISC','THE',FALSE));
           BASEFISC.Value := MONTANTHT.Value + fTVARecuperable - fTVARecuperee;  // BTY 03/06 FQ 17711;
           end;
        // Pour un composant remplaçant avec plan fiscal, bloquer la base à RC2
        if (fRemplace<> '') then
           begin
           TypeBaseFiscale.itemindex:=TypeBaseFiscale.Items.Indexof(RechDom('TITYPEBASEFISC','TR2',False));
           end;
        bDureeRestF.Checked := fbFuturVNF;
        bGestionFiscale.Enabled := False;
        bGestionFiscale.Checked := fbGestionfiscale;
        end
     else
        begin
        Typebasefiscale.itemindex:=typebasefiscale.items.indexof(RechDom('TITYPEBASEFISC','THE',FALSE));
        BASEFISC.Value := MONTANTHT.Value + fTVARecuperable - fTVARecuperee;  // BTY 03/06 FQ 17711;
        bGestionFiscale.Enabled := not (fbDPI); // ouvert si immo sans DPI
        bGestionFiscale.Checked := False;
        bDureeRestF.Checked := False;
        // Détermine case Gestion fiscale théorique
        DetermineGestionFiscale;
        // DPI ou prime ou subvention => pas de Gestion fiscale
        if (fSBVPRI<>0) or (fSBVMT<>0) or (fbDPI) then
           begin
           bGestionFiscale.Enabled := False;
           bGestionFiscale.Checked := False;
           end;
        end;
     end
  else
     begin
     typebasefiscale.visible:=false;
     tbasefisc.visible:=false;
     basefisc.visible:=false;
     BaseFiscaleEgale.Visible:=false;
     bDureeRestF.Visible := False;
     bDureeRestF.Checked := False;
     bGestionFiscale.Visible:= False;
     bGestionFiscale.Checked := False;
     end;
end;

//MVG 09/05

procedure TFModBas.TYPEBASEFISCALEChange(Sender: TObject);
begin
  inherited;
  if TYPEBASEFISCALE.Value='ECO' then
  begin
     basefisc.value := BaseEco.value;
  end
  else
  begin
     basefisc.value := MontantHT.value + fTVARecuperable - fTVARecuperee;  // BTY 03/06 FQ 17711;
  end;
  DetermineGestionFiscale;
end;
// MVG 09/05


procedure TFModBas.DetermineGestionFiscale;
var QImmo : TQuery;
    wPlan : TPlanAmort;
    wCumulEco,wCumulFiscal : double;
    CumulEcoTheorique : double ;
    CumulFiscTheorique : double ;

begin
   // Case bloquée si plan fiscal existant
   if bFiscal then
   begin
      bGestionFiscale.Enabled := False;
      bGestionFiscale.Checked := fbGestionfiscale;
   end
   else
   begin
      // DPI ou prime ou subvention => pas de Gestion fiscale
      if (fSBVPRI<>0) or (fSBVMT<>0) or (fbDPI) then
      begin
         bGestionFiscale.Enabled := False;
         bGestionFiscale.Checked := False;
      end
      else
      begin
         wPlan:=TPlanAmort.create(true) ;

         try
           QImmo := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+fCode+'"',False);
           if not QImmo.Eof then
           begin
             wPlan.Charge (QImmo);
             wPlan.Recupere(fCode,QImmo.FindField('I_PLANACTIF').AsString);
             wPlan.GetCumulsDotExercice(VHImmo^.Encours.Deb, wCumulEco,wCumulFiscal,True,True,true);
             wPlan.VncRestEco := bDureeRest.Checked;
             wPlan.AmortEco.ModifieMethodeAmort(wPlan.AmortEco.DateFinAmort,
                                                wPlan.AmortEco.Revision,
                                                BASEECO.Value,
                                                wPlan.AmortEco.Taux,
                                                wPlan.AmortEco.Duree,
                                                wPlan.AmortEco.Methode,
                                                wPlan.ImmoCloture);
             wPlan.CalculDotation(wPlan.AmortEco, VHImmo^.Encours.Deb, False);
             wPlan.VncRestFisc := bDureeRestF.Checked;
             wPlan.AmortFisc.ModifieMethodeAmort(wPlan.AmortEco.DateFinAmort,
                                                wPlan.AmortFisc.Revision,
                                                BASEFisc.Value,
                                                wPlan.AmortEco.Taux,
                                                wPlan.AmortEco.Duree,
                                                wPlan.AmortEco.Methode,
                                                wPlan.ImmoCloture);
             wPlan.CalculDotation(wPlan.AmortFisc, VHImmo^.Encours.Deb, False);
             // antérieurs THEORIQUES éco et fiscal selon les données saisies
             wPlan.CalculReprises;
             CumulEcoTheorique := wPlan.AmortEco.Reprise;
             CumulFiscTheorique := wPlan.AmortFisc.Reprise;

             QImmo.Edit;
             QImmo.FindField ('I_MONTANTHT').AsFloat := MONTANTHT.Value;
             QImmo.FindField ('I_BASEECO').AsFloat := BASEECO.Value;
             QImmo.FindField ('I_BASEFISC').AsFloat := BASEFISC.Value;
             QImmo.FindField ('I_METHODEFISC').AsString := QImmo.FindField ('I_METHODEECO').AsString;
             QImmo.FindField ('I_DUREEFISC').AsInteger := QImmo.FindField ('I_DUREEECO').AsInteger;
             QImmo.FindField ('I_TAUXFISC').AsFloat := QImmo.FindField ('I_TAUXECO').AsFloat;
             if bDureeRest.Checked then
               QImmo.FindField('I_JOURNALA').AsString := '***'
             else
               QImmo.FindField('I_JOURNALA').AsString := '';
             if bDureeRestF.Checked then
               QImmo.FindField('I_FUTURVNFISC').AsString := '***'
             else
               QImmo.FindField('I_FUTURVNFISC').AsString := '';

             // Antérieurs ECO > Fisc et <> 0 ou non conditionnent la valeur de "Gestion fiscale"
             if (wPlan.Determine_Gestion_Fiscale(QImmo, CumulEcoTheorique, CumulFiscTheorique))=True then
             begin
               bGestionFiscale.Checked := True;
               bGestionFiscale.Enabled := False;
             end else
             begin
               bGestionFiscale.Checked := False;
               bGestionFiscale.Enabled := True;
             end;
           end;
           Ferme (QImmo);
         finally
           wPlan.free ; //Detruit;
         end;
      end; // fSBVPRI
   end; // bfiscal
end;

end.
