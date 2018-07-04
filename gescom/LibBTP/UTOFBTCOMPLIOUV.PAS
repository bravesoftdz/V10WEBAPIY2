{***********UNITE*************************************************
Auteur  ...... : Lionel Santucci
Créé le ...... : 04/07/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : BTCOMPLIGNE ()
Mots clefs ... : TOF;BTCOMPLIGNE
*****************************************************************}
Unit UTOFBTCOMPLIOUV ;

Interface
uses
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     MajTable, Fe_Main ,
{$ENDIF}
      M3FP, StdCtrls,Controls,Classes,db,forms,sysutils,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,UTOF, AglInit,Agent,EntGC,saisutil,
      Dicobtp,HTB97,FactUtil,UAFO_Ressource,Vierge,UtilPGI,LookUp,graphics,
      utilgc,HrichOle,UtilArticle,ParamSoc,FactTvaMilliem;

Type
  TOF_BTCOMPLIOUV = Class (TOF)
  private
  LaRessource, CodeTiers,CodeAffaire : String;
  EnHt,Ouvrage : boolean;
  PA,PR,PuHtdev,PuTTCDev,Puht,puttc : THNumEdit;
  RatioTTC : THNumedit;
  TOBRepart,TOBLIgne,TOBPiece : TOB;
  COEFFG,COEFMARG,COEFFR,COEFFC : THNumEdit;
  APPLICCOEFFG,APPLICCOEFFC,APPLICCOEFFR : TcheckBox;
  BloquePv,ApplicFRST,ApplicFCST,PrestationST : boolean;
  BLOQUETARIF : TCheckBox;
  MontantPA,MontantPAFR,MontantPAFG,MontantPAFC,MontantFG,MontantFR,MontantFC,MontantPR : double;
  Qte : THnumEdit;
  FRFixed,FCFixed:   boolean;
  DEV : RDevise;
  TauxFg : double;
  procedure AffecteRessourceLigne;
  procedure ValideChamps;
    procedure calculeCoefs;
    procedure PuHtDevExit(Sender: TObject);
    procedure PuTTCDevExit(Sender: TObject);
    procedure Setdecimals;
    procedure SetComponents;
//    procedure COEFFGExit(Sender: TObject);
    procedure PAExit(Sender: TObject);
    procedure PRExit(Sender: TObject);
    procedure COEFMARGExit(Sender: TObject);
    procedure AppliqueCoefFGClick (SEnder : Tobject);
    procedure AppliqueCoefFRClick (SEnder : Tobject);
    procedure AppliqueCoefFCClick (SEnder : Tobject);
    procedure calculeFromDev(PuTTC: THNumEdit; ValueDev: double; DEV: Rdevise);
    procedure DeviseToPivot(ZOnePivot, ZoneDev: THnumEdit);
    procedure CalcPrLocal;
    procedure Getvalues;
    procedure COEFFGExit(Sender: TObject);
    procedure Setvalues;
    procedure COEFFRExit(Sender: TObject);
    procedure AppliqueCoefSousDetail(TOBCourante: TOB; Champ: string;
      Application: boolean);

  public
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

const
	TexteMsgCompl: array[1..7] of string 	= (
          {1}        'Vous devez renseigner un code tiers livré valide',
          {2}        'Vous devez renseigner un code tiers facturé valide',
          {3}        'Vous devez renseigner un code commercial valide',
          {4}        'Vous devez renseigner une ressource valide',
          {5}        'Vous devez renseigner un apporteur valide',
          {6}        'Changement de tiers interdit, l''affaire saisie n''est pas sur le tiers initial',
          {7}        'Code affaire non valide'
            );

Implementation

procedure TOF_BTCOMPLIOUV.AffecteRessourceLigne ;
Var Ressource : TAFO_Ressource;
    CodeRes : string;
BEGIN
CodeRes := GetControlText('BLO_RESSOURCE');
if (CodeRes ='') or (LaTOB = Nil) then exit;
if CodeRes = LaRessource then exit;

Ressource := TAFO_Ressource.Create(CodeRes);
// revoir l'appel de la fonction modifiée par affaire
//CoutPR := Ressource.PRRessource(Nil,LaTOB.GetValue('BLO_CODEARTICLE'),CodeAffaire);
//SetControlText ('BLO_PMRP',FloatToStr(CoutPR));
LaRessource := CodeRes;
Ressource.Free;
END;


procedure TOF_BTCOMPLIOUV.OnUpdate ;
begin
  Inherited ;
ValideChamps;
//Contrôle de la ressource saisie à la ligne
if (GetControlText('BLO_RESSOURCE') <> '' ) then
   if Not(LookupValueExist(THEdit(Getcontrol('BLO_RESSOURCE')))) then
      BEGIN
      PGIBoxAF(TexteMsgCompl[4],Ecran.Caption);
      SetFocusControl('BLO_RESSOURCE') ; TForm(Ecran).ModalResult:=0;
      Exit;
      END;
// Contrôle de l'apporteur d'affaire
if (GetControlText('BLO_APPORTEUR') <> '' ) then
   if Not(LookupValueExist(THEdit(Getcontrol('BLO_APPORTEUR')))) then
      BEGIN
      PGIBoxAF(TexteMsgCompl[5],Ecran.Caption);
      SetFocusControl('BLO_APPORTEUR') ; TForm(Ecran).ModalResult:=0;
      Exit;
      END;
// Contrôle du représentant
if (GetControlText('BLO_REPRESENTANT') <> '' ) then
   if Not(LookupValueExist(THEdit(Getcontrol('BLO_REPRESENTANT')))) then
      BEGIN
      PGIBoxAF(TexteMsgCompl[3],Ecran.Caption);
      SetFocusControl('BLO_REPRESENTANT') ;
      TForm(Ecran).ModalResult:=0;
      Exit;
      END;
LaTOB.putValue('BLO_COEFFG',CoefFg.value-1);
if  APPLICCOEFFG.checked then  LaTob.PutValue('BLO_NONAPPLICFG','-')
												 else LaTob.PutValue('BLO_NONAPPLICFG','X');
if  APPLICCOEFFR.checked then  LaTob.PutValue('BLO_NONAPPLICFRAIS','-')
												 else LaTob.PutValue('BLO_NONAPPLICFRAIS','X');
if  APPLICCOEFFC.checked then  LaTob.PutValue('BLO_NONAPPLICFC','-')
												 else LaTob.PutValue('BLO_NONAPPLICFC','X');
LaTOB.putValue('BLO_COEFMARG',CoefMarg.value);
LaTOB.PutValue('POURCENTMARG',Arrondi((LaTOB.GetValue('BLO_COEFMARG')-1)*100,2));
//TheTob:=LaTOB ;
end ;

procedure TOF_BTCOMPLIOUV.calculeCoefs;
var Montantmarg,PourcentMarg,MontantFrais : double;
begin
PourcentMarg := 0;
if Qte.value <> 0 then MontantFrais := (MontantFG + MontantFC + MontantFR)/Qte.value;
SetControlProperty ('MONTANTFRAIS','value',MontantFrais);
//
//
MontantMarg := PUHT.value - PR.value;
SetControlProperty ('MONTANTMARG','value',montantMarg);
//
if GetParamSocSecur('SO_BTGESTIONMARQ', False) then
begin
  if PUHT.value <> 0 then PourcentMarg := (MontantMarg / PUHT.value)*100 else PourcentMarg := 0;
end else
begin
  if PR.value <> 0 then PourcentMarg := (MontantMarg / PR.value)*100 else PourcentMarg := 0;
end;
SetControlProperty ('POURCENTMARG','value',PourcentMarg);
end;

procedure TOF_BTCOMPLIOUV.OnLoad ;
Var Nature: string;
begin
  Inherited ;
EnHT := (LaTOB.GetValue('BLO_FACTUREHT')='X');
ApplicFRST := (TOBPIece.GetValue('GP_APPLICFGST')='X');
ApplicFCST := (TOBPIece.GetValue('GP_APPLICFCST')='X');
if TOBPiece.getValue('GP_COEFFR')<> 0 then FRFixed := true else FRFixed := false;
if TOBPiece.getValue('GP_COEFFC')<> 0 then FCFixed := true else FCFixed := false;

PrestationST := IsPrestationST (LaTOb);
BloquePV := (TOBLIgne.GetValue('GL_BLOQUETARIF')='X');
BLOQUETARIF.checked := BloquePv;
DEV.Code:=laTob.getvalue('BLO_DEVISE') ;
GetInfosDevise(DEV) ;
DEV.Taux:=LaTob.getValue('BLO_TAUXDEV');
if LaTOB.getValue('BLO_PUHT') <> 0 then RATIOTTC.value := LaTOB.getValue('BLO_PUTTC')/LaTOB.getValue('BLO_PUHT')
                                   else RATIOTTC.value := 0;

if LaTOB.getValue('BLO_TYPEARTICLE') <> 'PRE' then Setcontrolvisible('GB_RESSOURCE',false);
LaRessource := LaTOB.GetValue('BLO_RESSOURCE');
Nature := LaTOB.GetValue('BLO_NATUREPIECEG');
Ouvrage :=  (LaTOB.GetValue('BLO_TYPEARTICLE')='OUV') OR (LaTOB.GetValue('BLO_TYPEARTICLE')='ARP');
if EnHT Then
begin
	PUTTCDEV.enabled := false;
	PUTTCDEV.Color := ClInactiveCaptionText;
	PUHTDEV.enabled := True;
	PUHTDEV.Color := ClWindow;
end else
begin
	PUTTCDEV.enabled := True;
	PUTTCDEV.Color := ClWindow;
	PUHTDEV.enabled := false;
	PUHTDEV.Color := ClInactiveCaptionText;
end;
if (Ouvrage) and (not isExistsArticle (trim(GetParamsoc('SO_BTECARTPMA')))) then
begin
	PUTTCDEV.Enabled := false;
	PUTTCDEV.color := ClInactiveCaptionText;
	PUHTDEV.Enabled := false;
	PUHTDEV.color := ClInactiveCaptionText;
end;

if BloquePv Then
begin
	PUHTDEV.Enabled := false;
	PUHTDEV.color := ClInactiveCaptionText;
	PUTTCDEV.Enabled := false;
	PUTTCDEV.color := ClInactiveCaptionText;
	COEFMARG.Enabled := false;
	COEFMARG.color := ClInactiveCaptionText;
end;

if (ouvrage) then
begin
  PA.enabled := false;
	PA.color := ClInactiveCaptionText;
  COEFFG.enabled := false;
	COEFFG.color := ClInactiveCaptionText;
end;

if ((FRFixed) and (PrestationST) and (ApplicFRST)) then
begin
  PR.enabled := false;
	PR.color := ClInactiveCaptionText;
end;

if ((FRFixed) and (not prestationST) and (APPLICCOEFFR.checked)) or (ouvrage) then
begin
  PR.enabled := false;
	PR.color := ClInactiveCaptionText;
end;

if ((FCFixed) and (not prestationST) and (APPLICCOEFFC.checked)) or (ouvrage) then
begin
  PR.enabled := false;
	PR.color := ClInactiveCaptionText;
end;

if ((FCFixed) or (FRFIxed) or (ouvrage)) then
begin
  PR.enabled := false;
	PR.color := ClInactiveCaptionText;
end;

if PuHTDEV.value > 0 then RatioTTC.Value := PuTTCDev.value / PuHtDEV.Value else RatioTTC.Value := 1;

Getvalues;
// Ne faire qu'une fois pour recup les coefs
CoeffG.value := LaTOB.GetValue('BLO_COEFFG')+1;
if (MontantPAFC <> 0) and (MontantFC <> 0) then CoeffC.value := (MontantFC+MontantFG+MontantPAFC)/(MontantPAFC+MontantFG)
																					 else CoeffC.value := 0;
if (MontantPAFR <> 0) and (MontantFR <> 0) then CoeffR.value := (MontantFG+MontantFR+MontantPAFR+MontantFC)/(MontantPAFR+MontantFC+MontantFG)
																					 else CoeffR.value := 0;
calculeCoefs;

end ;

procedure TOF_BTCOMPLIOUV.Setdecimals;
begin
	THNumEdit(GetControl('BLO_QTEFACT')).NumericType := ntDecimal;
	THNumEdit(GetControl('BLO_QTEFACT')).Decimals := V_PGI.OkdecQ;
	THNumEdit(GetControl('BLO_DPA')).NumericType := ntDecimal;
	THNumEdit(GetControl('BLO_DPA')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('BLO_DPR')).NumericType := ntDecimal;
	THNumEdit(GetControl('BLO_DPR')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('MONTANTFRAIS')).NumericType := ntDecimal;
	THNumEdit(GetControl('MONTANTFRAIS')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('BLO_PUHTBASE')).NumericType := ntDecimal;
	THNumEdit(GetControl('BLO_PUHTBASE')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('MONTANTMARG')).NumericType := ntDecimal;
	THNumEdit(GetControl('MONTANTMARG')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('BLO_PUHTDEV')).NumericType := ntDecimal;
	THNumEdit(GetControl('BLO_PUHTDEV')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('BLO_PUTTCDEV')).NumericType := ntDecimal;
	THNumEdit(GetControl('BLO_PUTTCDEV')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('BLO_PUHT')).NumericType := ntDecimal;
	THNumEdit(GetControl('BLO_PUHT')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('BLO_PUTTC')).NumericType := ntDecimal;
	THNumEdit(GetControl('BLO_PUTTC')).Decimals := V_PGI.OkdecP;
end;

procedure TOF_BTCOMPLIOUV.OnArgument (S : String ) ;
var Critere,ChampMul,ValMul : string;
x : Integer;
begin
  Inherited ;
//
Repeat
    Critere:=uppercase(Trim(ReadTokenSt(S))) ;
    if Critere<>'' then
    begin
        x:=pos('=',Critere);
        if x<>0 then
        begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));

           if ChampMul='AFFAIRE' then CodeAffaire:=ValMul ;
           if ChampMul='TIERS' then CodeTiers := ValMul;
        end;
    end;
until  Critere='';
//
TOBLIgne := TOB(LaTOB.data);
TOBPiece := TOBLigne.parent;
TOBREpart := TOB(TOBLigne.data);
SetcontrolVisible('TIERSAFFAIRE', false);
SetcontrolVisible('LIBTIERSAFFAIRE', false);
SetcontrolVisible('BEFFACEAFF1', false);
//if Not(bVente) then SetControlText('TBLO_TIERS','&Fournisseur');
GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'BLO_LIBREART', 10, '');

Qte := THNumEdit(GetCOntrol('BLO_QTEFACT'));
PA := THNumEdit(GetControl ('BLO_DPA'));
PA.OnExit := PAExit;
PR := THNumEdit(GetControl ('BLO_DPR'));
PR.OnExit := PRExit;
PuHTDEV := THNumEdit(GetControl ('BLO_PUHTDEV'));
PuHtDEV.OnExit := PuHTDevExit;
PuHT := THNumEdit(GetControl ('BLO_PUHT'));
PuTTCDev := THNumEdit(GetCOntrol('BLO_PUTTCDEV'));
PuTTCdev.OnExit := PuTTCDevExit;
PuTTC := THNumEdit(GetCOntrol('BLO_PUTTC'));
RATIOTTC := THNumEdit(GetControl ('RATIOTTC'));

//
COEFFG := THNumEdit(GetCOntrol('COEFFG')); COEFFG.onExit  := CoeffGExit;
COEFFC := THNumEdit(GetCOntrol('COEFFC'));
COEFFR := THNumEdit(GetCOntrol('COEFFR'));
COEFMARG := THNumEdit(GetCOntrol('COEFMARG'));
COEFMARG.value := LaTOB.GetValue('BLO_COEFMARG');
COEFMARG.OnExit := COEFMARGExit;
//
APPLICCOEFFG := TCheckBox(GetControl('APPLICCOEFFG'));
APPLICCOEFFC := TCheckBox(GetControl('APPLICCOEFFC'));
APPLICCOEFFR := TCheckBox(GetControl('APPLICCOEFFR'));
BLOQUETARIF := TCheckBox(GetControl('BLOQUETARIF'));

Setdecimals;
SetComponents;
//uniquement en line
//TCheckBox(GetCOntrol('APPLICCOEFFR')).visible := false;
//TGroupBox(GetCOntrol('G_COMM')).visible := false;
//
APPLICCOEFFG.OnClick := AppliqueCoefFGClick;
APPLICCOEFFC.OnClick := AppliqueCoefFCClick;
APPLICCOEFFR.OnClick := AppliqueCoefFRClick;
	if TOBLIgne.GetString('GL_TYPENOMENC')='OU1' then
  begin
    SetControlEnabled ('BLO_FAMILLETAXE1',False);
    SetControltext    ('MILLI1','Comptabilisation au niveau ouvrage');
    SetControlVisible ('MILLI1',True);
    SetControlEnabled ('BLO_FAMILLETAXE2',False);
    SetControlEnabled ('BLO_FAMILLETAXE3',False);
    SetControlEnabled ('BLO_FAMILLETAXE4',False);
    SetControlEnabled ('BLO_FAMILLETAXE5',False);
  end;


end ;

procedure TOF_BTCOMPLIOUV.PAExit (Sender : TObject);
begin
  MontantPA := PA.Value * QTe.Value;
  if not Ouvrage then
  begin
    if APPLICCOEFFC.Checked then MontantPAFC := MontantPA else MontantPA := 0;
    if APPLICCOEFFR.Checked then MontantPAFR := MontantPA else MontantPA := 0;
  end;
  CalcPrLocal;
  CalculeCoefs;
end;

procedure TOF_BTCOMPLIOUV.PRExit (Sender : TObject);
begin
	if PA.value <> 0 then COEFFG.Value := MontantFG/MontantPAFG;
  CalcPrLocal;
  CalculeCoefs;
end;

procedure TOF_BTCOMPLIOUV.COEFFRExit (Sender : TObject);
begin
  CalcPrLocal;
  calculeCoefs;
end;


procedure TOF_BTCOMPLIOUV.COEFMARGExit (Sender : TObject);
begin
  CalcPrLocal;
  CalculeCoefs;
end;


procedure TOF_BTCOMPLIOUV.PuHtDevExit (Sender : TObject);
var savPb : boolean;
begin
  DeviseToPivot ( PUHT,PUHtDEv) ;
  SavPb := BloquePV;
  BloquePV := true;
  CalcPrLocal;
  BloquePv := SavPb;
  CalculeCoefs;
end;

procedure TOF_BTCOMPLIOUV.OnClose ;
begin
  Inherited ;
end ;

Procedure AGLAffecteRessourceLigne ( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
BEGIN
LaTof := nil;
F:=TForm(Longint(Parms[0]));
if (F is TFVierge) then LaTof:=TFVierge(F).Latof;
if (LaTof is TOF_BTCOMPLIOUV) then TOF_BTCOMPLIOUV(LaTof).AffecteRessourceLigne;
END;

procedure TOF_BTCOMPLIOUV.ValideChamps;
begin
NextControl (Ecran,true);
end;

procedure TOF_BTCOMPLIOUV.PuTTCDevExit(Sender: TObject);
var SavPb : boolean;
begin
	DeviseToPivot (PuTTC,PuTTCDev);
  if RatioTTc.value <> 0 then
  begin
    PUHTDEV.Value := Arrondi (PuTTCDev.Value/RATIOTTC.Value,V_PGI.OkDecP);
  end;
  DEVISETOPIVOT (PuHt,PuHtDev);
  SavPb := BloquePv;
  BloquePV := true;
  CalcPrLocal;
  bloquePv := SavPb;
  CalculeCoefs;
end;

procedure TOF_BTCOMPLIOUV.SetComponents;
var Indice : integer;
		UneTaxe : TOB;
begin
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    Unetaxe  := VH_GC.TOBParamTaxe.detail[Indice];
    // non paramétré
    THLabel (GetControl('TBLO_FAMILLETAXE'+inttostr(Indice+1))).Visible := false;
    THValComboBox (GetControl('BLO_FAMILLETAXE'+inttostr(Indice+1))).Visible := false;
    THNumEdit(GetCOntrol('MILLIEME'+InttoStr(indice+1))).Visible := false;
    THLabel(GetCOntrol('MILLI'+InttoStr(indice+1))).Visible := false;
    if UneTaxe.getValue('BPT_TYPETAXE') <> '' then
    begin
      if UneTaxe.getValue('BPT_TYPETAXE') = 'TVA' Then
      begin
      //
      	if (Indice > 0) and ((TOBRepart = nil)  or (TOBRepart.detail.count = 0 )) then continue;
      //
        THValComboBox (GetControl('BLO_FAMILLETAXE'+inttostr(Indice+1))).DataType := 'TTTVA';
        THValComboBox (GetControl('BLO_FAMILLETAXE'+inttostr(Indice+1))).Visible := True;
        THLabel (GetControl('TBLO_FAMILLETAXE'+inttostr(Indice+1))).Caption := UneTaxe.getValue('BPT_LIBELLE');
        THLabel (GetControl('TBLO_FAMILLETAXE'+inttostr(Indice+1))).visible := true;
        if (TOBRepart <> nil)  and (TOBRepart.detail.count > 0 ) then
        begin
          THNumEdit(GetCOntrol('MILLIEME'+InttoStr(indice+1))).Value :=  GetMillieme('TX'+IntToStr(Indice+1),TOBRepart);
          if THNumEdit(GetCOntrol('MILLIEME'+InttoStr(indice+1))).Value > 0 then
          begin
            THNumEdit(GetCOntrol('MILLIEME'+InttoStr(indice+1))).Visible := true;
            THLabel(GetCOntrol('MILLI'+InttoStr(indice+1))).Visible := true;
          end;
        end;
      end else
      begin if UneTaxe.getValue('BPT_TYPETAXE') = 'TPF' then
      	THValComboBox (GetControl('BLO_FAMILLETAXE'+inttostr(Indice+1))).DataType := 'TTTPF';
      	THValComboBox (GetControl('BLO_FAMILLETAXE'+inttostr(Indice+1))).Visible := True;// {$IFDEF LINE}false{$ENDIF};
      	THLabel (GetControl('TBLO_FAMILLETAXE'+inttostr(Indice+1))).Caption := UneTaxe.getValue('BPT_LIBELLE');
      	THLabel (GetControl('TBLO_FAMILLETAXE'+inttostr(Indice+1))).visible := True; {$IFDEF LINE}false{$ENDIF};
      end;
      if (TOBRepart <> nil) and (TOBRepart.detail.count > 0) then
         THValComboBox (GetControl('BLO_FAMILLETAXE'+inttostr(Indice+1))).Enabled := false;
    end;
  end;
  APPLICCOEFFR.checked := not (LaTob.getValue('BLO_NONAPPLICFRAIS')='X');
  APPLICCOEFFC.checked := not (LaTob.getValue('BLO_NONAPPLICFC')='X');
  APPLICCOEFFG.checked := not (LaTob.getValue('BLO_NONAPPLICFG')='X');
end;

procedure TOF_BTCOMPLIOUV.AppliqueCoefFRClick(SEnder: Tobject);
begin
  if LaTOB.Detail.count > 0 then
  begin
    AppliqueCoefSousDetail (laTOB,'NONAPPLICFRAIS',APPLICCOEFFR.checked);
  end;
  if not APPLICCOEFFR.checked then
  begin
  	MontantFR := 0;
  end else
  begin
  	if TOBPiece.getValue('GP_COEFFR') <> 0 then
    begin
    	if Ouvrage then
      begin
      	MontantPAFR := LaTOB.GetValue('BLO_MONTANTPAFR');
      end else
      begin
      	MontantPAFR := LaTOB.GetValue('BLO_MONTANTPA');
      end;
    	if ((PrestationST) and (ApplicFRSt)) or (not PrestationST) then
      begin
        MontantFR := (MontantPAFR + MontantFC + MontantFG) * TOBPiece.getValue('GP_COEFFR');
        if MontantPAFG <> 0 then
        begin
          COEFFR.Value := (MontantPAFR+MontantFC+MontantFR+MontantFG) / (MontantPAFR+MontantFc+MontantFR);
        end;
      end else
      begin
        MontantFR := 0;
        COEFFR.value := 0;
      end;
    end;
  end;
  CalcPrLocal;
  calculeCoefs;
end;

procedure TOF_BTCOMPLIOUV.AppliqueCoefFCClick(SEnder: Tobject);
begin
  if LaTOB.Detail.count > 0 then
  begin
    AppliqueCoefSousDetail (laTOB,'NONAPPLICFC',APPLICCOEFFG.checked);
  end;
  if not APPLICCOEFFC.checked then
  begin
  	MontantFC := 0;
  end else
  begin
  	if TOBPiece.getValue('GP_COEFFC') <> 0 then
    begin
    	if Ouvrage Then
      begin
    		MontantPAFC := LaTOB.GetValue('BLO_MONTANTPAFC');
      end else
      begin
    		MontantPAFC := LaTOB.GetValue('BLO_MONTANTPA');
      end;
    	if ((PrestationST) and (ApplicFcSt)) or (not PrestationST) then
      begin
        MontantFC := (MontantPAFC+MontantFG) * TOBPiece.getValue('GP_COEFFC');
        if MontantPAFC <> 0 then
        begin
          COEFFC.Value := (MontantPAFC+MontantFC+MontantFG) / (MontantPAFC+MontantFG);
        end;
      end else
      begin
        MontantFC := 0;
        COefFC.value := 0;
      end;
    end;
  end;
  CalcPrLocal;
  calculeCoefs;
end;

procedure TOF_BTCOMPLIOUV.calculeFromDev(PuTTC: THNumEdit;
  ValueDev: double; DEV: Rdevise);
begin
  if DEV.Code <> V_PGI.DevisePivot  then
  begin
      PuTTC.Value := DEVISETOEURO (ValueDev,DEV.Taux,DEV.quotite);
  end else
  begin
      PuTTC.Value := ValueDev;
  end;
end;

procedure TOF_BTCOMPLIOUV.DeviseToPivot(ZOnePivot, ZoneDev: THnumEdit);
begin
  if DEV.Code <> V_PGI.DevisePivot  then
  begin
      ZOnePivot.Value := DEVISETOEURO (ZoneDev.Value,DEV.Taux,DEV.quotite);
  end else
  begin
      ZOnePivot.Value := ZoneDev.Value;
  end;
end;


procedure TOF_BTCOMPLIOUV.CalcPrLocal;
begin
  if APPLICCOEFFG.checked then
  BEGIN
    MontantFG := (MontantPAFG * COEFFG.Value) - MontantPAFG;
  END else
  BEGIN
    MontantFG := 0;
  END;

	if APPLICCOEFFC.checked then
  begin
  	if COEFFC.Value  <> 0 then MontantFC := ((MontantPAFC+MontantFG) * COEFFC.Value) - (MontantPAFC+MontantFG)
    											else MontantFC := 0;
  end else
  begin
  	MontantFC := 0;
  end;

	if APPLICCOEFFR.checked then
  begin
  	if COEFFR.Value <> 0 then montantFR := ((MontantPAFR + MontantFC + MontantFG) * COEFFR.Value) - (MontantPAFR + MontantFC+ MontantFG)
    									   else MontantFR := 0;
  end else
  begin
  	montantFR := 0;
  end;

  if MontantFG+MontantFR+MontantFC <> 0 then
  begin
  	MontantPR := MontantPA+MontantFG+MontantFC+MontantFR;
    if Qte.value <> 0 then PR.Value := arrondi(MontantPR / Qte.value,V_PGI.OKDecp);
  end else
  begin
  	MontantPR := MontantPA;
    if Qte.value <> 0 then PR.Value := arrondi(MontantPR / Qte.value,V_PGI.OKDecp);
  end;

	if not BloquePV Then
  begin
    PUHt.value := Arrondi (PR.Value * COEFMARG.Value,V_PGI.OkDecP);
    if DEV.Code <> V_PGI.DevisePivot  then
    begin
      PUHtDEV.Value := EUROTODEVISE (PUHT.value,DEV.Taux,DEV.Quotite,V_PGI.OkDecP);
    end else
    begin
     PUHtDEV.Value := PUHT.Value;
    end;
    PUTTCDEV.Value := Arrondi (PuHTDev.Value*RATIOTTC.Value,V_PGI.OkDecP);
    PUTTC.Value := Arrondi (PuHT.Value*RATIOTTC.Value,V_PGI.OkDecP);
  end else
  begin
  	COEFMARG.Value := arrondi (PUHt.value/Pr.value,4);
  end;

end;


procedure TOF_BTCOMPLIOUV.Getvalues;
begin
	MontantPA := LaTOB.GetValue('BLO_MONTANTPA');
	MontantPAFG := LaTOB.GetValue('BLO_MONTANTPAFG');
  MontantFG := LaTOB.GetValue('BLO_MONTANTFG');
	MontantPAFC := LaTOB.GetValue('BLO_MONTANTPAFC');
  MontantFC := LaTOB.GetValue('BLO_MONTANTFC');
  MontantPAFR := LaTOB.GetValue('BLO_MONTANTPAFR');
  montantFR := LaTOB.GetValue('BLO_MONTANTFR');
  MontantPR := LaTOB.GetValue('BLO_MONTANTPR')
end;


procedure TOF_BTCOMPLIOUV.COEFFGExit(Sender: TObject);
begin
  CalcPrLocal;
  calculeCoefs;
end;


procedure TOF_BTCOMPLIOUV.Setvalues;
begin
  if not Ouvrage then
  begin
		LaTOB.PutValue('BLO_MONTANTPA',MontantPA);
    LaTOB.PutValue('BLO_MONTANTPAFG',MontantPaFG);
    LaTOB.PutValue('BLO_MONTANTPAFC',MontantPaFC);
    LaTOB.PutValue('BLO_MONTANTPAFR',MontantPaFR);
  end;
  LaTOB.PutValue('BLO_MONTANTFG',MontantFG);
  LaTOB.PutValue('BLO_MONTANTFC',MontantFC);
  LaTOB.PutValue('BLO_MONTANTFR',MontantFR);
  if (MontantPAFC+MontantFG) <> 0 then LaTOB.PutValue('BLO_COEFFC',arrondi(MontantFC/(MontantPAFC+MontantFG),4));
  if (MontantPAFR+MontantFC+MontantFG) <> 0 then LaTOB.PutValue('BLO_COEFFR',arrondi(MontantFR/(MontantPAFR+MontantFC+MontantFG),4));
  LaTOB.PutValue('BLO_MONTANTPR',MontantPR);
end;

procedure TOF_BTCOMPLIOUV.AppliqueCoefFGClick(SEnder: Tobject);
begin
  if LaTOB.Detail.count > 0 then
  begin
    AppliqueCoefSousDetail (laTOB,'NONAPPLICFG',APPLICCOEFFG.checked);
  end;
  if not APPLICCOEFFG.checked then
  begin
  	MontantFG := 0;
  end else
  begin
    if Ouvrage then
    begin
      MontantPAFG := LaTOB.GetValue('BLO_MONTANTPAFG');
    end else
    begin
      MontantPAFG := LaTOB.GetValue('BLO_MONTANTPA');
    end;
    MontantFG := MontantPaFG * COEFFG.Value;
  end;
  CalcPrLocal;
  calculeCoefs;
end;

procedure TOF_BTCOMPLIOUV.AppliqueCoefSousDetail (TOBCourante: TOB; Champ : string; Application : boolean);
var Indice : integer;
    TOBC : TOB;
begin
  for Indice := 0 to TOBCourante.detail.count -1 do
  begin
    TOBC := TOBCourante.detail[Indice];
    if Application then TOBC.putValue('BLO_'+Champ,'-')
                   else TOBC.putValue('BLO_'+Champ,'X');
    if TOBC.detail.count > 0 then AppliqueCoefSousDetail (TOBC,Champ,Application);
  end;
end;

Initialization
  RegisterAglProc( 'ChangeRessourceLiOuv', True ,1, AGLAffecteRessourceLigne);
  registerclasses ( [ TOF_BTCOMPLIOUV ] ) ;
end.
