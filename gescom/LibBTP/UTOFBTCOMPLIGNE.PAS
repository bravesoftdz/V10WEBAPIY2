{***********UNITE*************************************************
Auteur  ...... : Lionel Santucci
Créé le ...... : 04/07/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : BTCOMPLIGNE ()
Mots clefs ... : TOF;BTCOMPLIGNE
*****************************************************************}
Unit UTOFBTCOMPLIGNE ;

Interface

uses {$IFDEF VER150} variants,{$ENDIF}
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     MajTable, Fe_Main ,
{$ENDIF}
      graphics,
      M3FP, StdCtrls,Controls,Classes,db,forms,sysutils,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,UTOF, AglInit,Agent,EntGC,saisutil,
      AffaireUtil,Dicobtp,HTB97,FactUtil,UAFO_Ressource,Vierge,UtilPGI,LookUp,
      UTOFBTCODEAFFAIRE,UTOFAFTRADUCCHAMPLIBRE,utilgc,HrichOle,UtilArticle,ParamSoc,FactTvaMilliem,
      SelectPhase,UtilPhases, AppelsUtil;

Type
  TOF_BTCOMPLIGNE = Class (TOF_BTCODEAFFAIRE)
  private
  BRAZAFFAIRE  : TToolbarButton97;
  PuDEVModifie : boolean;
  TOBPiece,TOBRepart : TOB;
  EnHt : boolean;
  FRFixed,FCFixed:   boolean;
  LaRessource, CodeTiers, AncAffaire : String;
  PieceAffaire, bVente : Boolean;
  PuHtDEv,PA,PR,PUHt,PuTTC,PuTTCDEV : THNumEdit;
  COEFFR,COEFMARG,RATIOTTC : THNumEdit;
  COEFFC,COEFFG : THNumEdit;
  APPLICCOEFFR,APPLICCOEFFC ,APPLICCOEFFG: TCheckBox;
  Qte : THnumEdit;
  DEV : RDevise;
  Ouvrage,PrestationST,ApplicFRST,ApplicFCST,BloquePV : boolean;
  TypeVentilOuv,TypeArt : string;
  BLOQUETARIF : TCheckBox;
  MontantPA,MontantPAFR,MontantPAFG,MontantPAFC,MontantFG,MontantFR,MontantFC,MontantPR : double;
  IsDejaFacture : boolean;
  procedure AffecteRessourceLigne;
  procedure ValideChamps;
    procedure calculeCoefs;
    procedure PuHtDevExit(Sender: TObject);
    procedure PuTTCDevExit(Sender: TObject);
    procedure PAExit(Sender: TObject);
    procedure PRExit(Sender: TObject);
    procedure COEFFRExit(Sender: TObject);
    procedure COEFFGExit (SENDER : TOBject);
    procedure COEFMARGExit(Sender: TObject);
    procedure SetDecimals;
    procedure ModeConsultPrix;
    procedure SetComponents;
		procedure BrechPhaseClick (Sender : Tobject);
    procedure PhaseChange(Sender: TObject);
    procedure BloqueTarifClick(Sender: Tobject);
    procedure CalcPrLocal;
    procedure Getvalues;
    procedure APPLICCOEFFGClick(Sender: Tobject);
    procedure APPLICCOEFFCClick(Sender: Tobject);
    procedure APPLICCOEFFRClick(Sender: Tobject);
    //procedure Setvalues;
    procedure DeviseToPivot(ZOnePivot, ZoneDev: THnumEdit);
    procedure BRazAffaireClick(Sender: TObject);
    procedure GESTIONPTCClick (Sender : TObject);
  public
    procedure bSelectAff1Click(Sender: TObject); override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

const
	TexteMsgCompl: array[1..8] of string 	= (
          {1}        'Vous devez renseigner un code tiers livré valide',
          {2}        'Vous devez renseigner un code tiers facturé valide',
          {3}        'Vous devez renseigner un code commercial valide',
          {4}        'Vous devez renseigner une ressource valide',
          {5}        'Vous devez renseigner un apporteur valide',
          {6}        'Changement de tiers interdit, l''affaire saisie n''est pas sur le tiers initial',
          {7}        'Code affaire non valide !',
          {8}        'Saisie non autorisée. Ce chantier n''est pas accepté.'
            );

Implementation
uses  BTPUtil,
      CalcOLEGenericBTP,
      UTofAfBaseCodeAffaire,
      FactureBTP;

procedure TOF_BTCOMPLIGNE.AffecteRessourceLigne ;
Var Ressource : TAFO_Ressource;
    CodeRes : string;
    CoutPR : double;
BEGIN
CoutPR := 0;
CodeRes := GetControlText('GL_RESSOURCE');
if (CodeRes ='') or (LaTOB = Nil) then exit;
if CodeRes = LaRessource then exit;

Ressource := TAFO_Ressource.Create(CodeRes);
// revoir l'appel de la fonction modifiée par affaire
//CoutPR := Ressource.PRRessource(Nil,LaTOB.GetValue('GL_CODEARTICLE'),LaTOB.GetValue('GL_AFFAIRE'));
SetControlText ('GL_PMRP',FloatToStr(CoutPR));
LaRessource := CodeRes;
Ressource.Free;
END;


procedure TOF_BTCOMPLIGNE.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
// ***Fait avant le on_argument (on_argument de la tof ancêtre ) **
if LaTob <> nil then
   begin
   if LaTob.FieldExists ('GL_NATUREPIECEG') then
      begin
      if GetInfoParPiece(LaTob.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT')='ACH' then
         bVente := False else bVente := True;
      end;
   end;
///******************

Aff:=THEdit(GetControl('GL_AFFAIRE'));
Aff0:=THEdit(GetControl('GL_AFFAIRE0'));
Aff1:=THEdit(GetControl('GL_AFFAIRE1'));
Aff2:=THEdit(GetControl('GL_AFFAIRE2')); Aff3:=THEdit(GetControl('GL_AFFAIRE3'));
Aff4:=THEdit(GetControl('GL_AVENANT'));
if bVente then Tiers := THEdit (GetControl('GL_TIERS'))
          else Tiers := THEdit (GetControl('TIERSAFFAIRE'));
end;

procedure TOF_BTCOMPLIGNE.OnUpdate ;
Var Aff0,Aff,Aff1,Aff2,Aff3,Aff4,Tiers : THEdit;
    Part0,Part1,Part2,Part3,part4,Domaine : string;
    NbAff : integer;
    EtatAffaire : String;
    TypeAffaire : String;
begin
  Inherited ;
  ValideChamps;
  //Contrôle de la ressource saisie à la ligne
  if (GetControlText('GL_RESSOURCE') <> '' ) then
   if Not(LookupValueExist(THEdit(Getcontrol('GL_RESSOURCE')))) then
      BEGIN
      PGIBoxAF(TexteMsgCompl[4],Ecran.Caption);
      SetFocusControl('GL_RESSOURCE') ; TForm(Ecran).ModalResult:=0;
      Exit;
      END;
  // Contrôle de l'apporteur d'affaire
  if (GetControlText('GL_APPORTEUR') <> '' ) then
   if Not(LookupValueExist(THEdit(Getcontrol('GL_APPORTEUR')))) then
      BEGIN
      PGIBoxAF(TexteMsgCompl[5],Ecran.Caption);
      SetFocusControl('GL_APPORTEUR') ; TForm(Ecran).ModalResult:=0;
      Exit;
      END;
  // Contrôle du représentant
  if (GetControlText('GL_REPRESENTANT') <> '' ) then
  // Correction fiche qualite N°  10113 --
    if not ExisteSQL('SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_COMMERCIAL="' +
                      GetControlText('GL_REPRESENTANT') + '"') then
    begin
      PGIBoxAF(TexteMsgCompl[3], Ecran.Caption);
      SetFocusControl('GL_REPRESENTANT');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  //
(*
   if Not(LookupValueExist(THEdit(Getcontrol('GL_REPRESENTANT')))) then
      BEGIN
      PGIBoxAF(TexteMsgCompl[3],Ecran.Caption);
      SetFocusControl('GL_REPRESENTANT') ;
      TForm(Ecran).ModalResult:=0;
      Exit;
      END;
   *)
  if not PieceAffaire then
  BEGIN
    // Contrôle de l'affaire
    Aff0:=THEdit(GetControl('GL_AFFAIRE0'));
    Aff:=THEdit(GetControl('GL_AFFAIRE'));
    Aff1:=THEdit(GetControl('GL_AFFAIRE1'));
    Aff2:=THEdit(GetControl('GL_AFFAIRE2'));
    Aff3:=THEdit(GetControl('GL_AFFAIRE3'));
    Aff4:=THEdit(GetControl('GL_AVENANT'));
    //
    Part0:=Aff0.text;
    Part1:=Aff1.text;
    Part2:=Aff2.text;
    Part3:=Aff3.text;
    Part4:=Aff4.text;
    //
    TypeAffaire := Aff0.text;

		Aff.text := '';
    if Part1 <> '' then
    begin
    	if Part0 = 'W' then
        Aff.text := RegroupePartiesAppel (Part0, Part1, PArt2, Part3, Part4)
      else
      	Aff.text := CodeAffaireRegroupe (Part0, Part1, PArt2, Part3, Part4, taModif, false, false, false);
    end;

    if bVente then
      Tiers:=THEdit(GetControl('GL_TIERS'))
    else
      Tiers:=THEdit(GetControl('TIERSAFFAIRE')) ;

    // controle du tiers associé à l'affaire
    if (Codetiers <>  Tiers.Text) And (bVente) then
    BEGIN
      SetFocusControl('GL_AFFAIRE1') ;
      PGIBoxAF(TexteMsgCompl[6],Ecran.Caption);
      SetFocusControl('GL_AFFAIRE1') ; TForm(Ecran).ModalResult:=0;
      Exit;
    END;

    if (Aff1.Text='') and (Aff2.Text='') and (Aff3.Text='') then
      NbAff := -1    // pas de test sur l'affaire
    else
      NbAff:=TestCleAffaire(Aff,Aff1,Aff2,Aff3,Aff4,TIERS,'A',True,False,False);

    if NbAff = 0 then
    BEGIN
      PGIBoxAF(TexteMsgCompl[7],Ecran.Caption);
      SetFocusControl('GL_AFFAIRE1') ; TForm(Ecran).ModalResult:=0;
      Exit;
    END;

    if NbAff <> -1 then
    begin
      if GetControlText('GP_VENTEACHAT')= 'ACH' then
      begin
        //Mis en commentaire par BRL le 5/05/2014 à la demande de DELABOUDINIERE
        //plus de contrôle pour les appels
        //If ((Part0 = 'A' ) AND (EtatAffaire <> 'ACP')) Or
        //   ((Part0 = 'W' ) AND (EtatAffaire = 'ECO')) then
        //FV1 : 10/06/2014 - FS#921 - DELABOUDINIERE : Revoir les contrôles sur appels et contrats en fonction du code état
        if not ControleAffaire(Aff.text, Ecran.caption,'ACH') then
        begin
          Aff.text := AncAffaire;
          BTPCodeAffaireDecoupe(AncAffaire,Part0, Part1,Part2,Part3,Part4,TaModif,False);
          THEdit(GetControl('GL_AFFAIRE0')).Text  := Part0;
          THEdit(GetControl('GL_AFFAIRE1')).Text  := Part1;
          THEdit(GetControl('GL_AFFAIRE2')).Text  := Part2;
          THEdit(GetControl('GL_AFFAIRE3')).Text  := Part3;
          THEdit(GetControl('GL_AVENANT')).Text   := Part4;
          SetFocusControl('GL_AFFAIRE1') ;
          LaTOB.putValue('GL_AFFAIRE',Aff.text);
          TForm(Ecran).ModalResult  :=0;
        end;
        (*
        else
        begin
          //FV1 : 18/09/2013 - FS#665 - En saisie de pièce achats, le contrôle sur affaire acceptée n'est pas fait dans le complément ligne
          if not (GetParamSocSecur('SO_BTINTERDIREACHATS',False)) Then
          begin
            EtatAffaire := GetChampsAffaire (Aff.text,'AFF_ETATAFFAIRE');
            if EtatAffaire='ENC' then
            begin
              if TypeAffaire = 'A' then
              begin
                if (not (LaTob.getString('GL_NATUREPIECEG')='BBO')) then PgiBox (TraduireMemoire('Attention, ce chantier n''est pas accepté'),TraduireMemoire('Contrôle chantier'));
              end;
            end;
          end;
        end;
        *)
      end;
    end;

    //FV1 : 26/09/2013 -  chargement du Code Affaire dans laTOB pour gestion changement de d'Affaire en modification
    LaTOB.putValue('GL_AFFAIRE',Aff.text);

    if AncAffaire <> LaTOB.GetValue('GL_AFFAIRE') then
    begin
      if GetParamSoc ('SO_BTLIVCHANTIER') then
      begin
        if (LaTOB.GetString('GL_NATUREPIECEG')<>'CBT') or
           ((LaTOB.GetString('GL_NATUREPIECEG')='CBT') and (not GetParamSocSecur('SO_BTLIVBESOINDEF',false))) then
        begin
          TcheckBox(getControl('LIVRAISONFOUR')).checked := true;
        end;
      end;
      Domaine := GetChampsAffaire (LaTOB.GetValue('GL_AFFAIRE'), 'AFF_DOMAINE');
      if Domaine <> '' then LaTOB.putValue('GL_DOMAINE',Domaine);
    end;
  END;

  if TcheckBox(getControl('LIVRAISONFOUR')).checked then LaTOB.putValue('GL_IDENTIFIANTWOL',-1)
                                                    else LaTOB.putValue('GL_IDENTIFIANTWOL',0);

  if ((Ouvrage) AND
     ((GetControlText('FAMILLETAXE1') <> '') OR (GetControlText('FAMILLETAXE2') <> '') OR
      (GetControlText('FAMILLETAXE3') <> '') OR (GetControlText('FAMILLETAXE4') <> '') OR
      (GetControlText('FAMILLETAXE5') <> ''))) or (not Ouvrage) then
  begin
    LaTOB.putValue('GL_FAMILLETAXE1',GetControlText('FAMILLETAXE1'));
    LaTOB.putValue('GL_FAMILLETAXE2',GetControlText('FAMILLETAXE2'));
    LaTOB.putValue('GL_FAMILLETAXE3',GetControlText('FAMILLETAXE3'));
    LaTOB.putValue('GL_FAMILLETAXE4',GetControlText('FAMILLETAXE4'));
    LaTOB.putValue('GL_FAMILLETAXE5',GetControlText('FAMILLETAXE5'));
  end;

  if IsBLobVide (ecran,LaTob,'GL_BLOCNOTE') then laTOB.putvalue('GL_BLOCNOTE','');
  //
if  TCheckBox(GetControl('APPLICCOEFFR')).checked then  LaTob.PutValue('GLC_NONAPPLICFRAIS','-')
																									else LaTob.PutValue('GLC_NONAPPLICFRAIS','X');
if  TCheckBox(GetControl('APPLICCOEFFG')).checked then  LaTob.PutValue('GLC_NONAPPLICFG','-')
																									else LaTob.PutValue('GLC_NONAPPLICFG','X');
if  TCheckBox(GetControl('APPLICCOEFFC')).checked then  LaTob.PutValue('GLC_NONAPPLICFC','-')
																									else LaTob.PutValue('GLC_NONAPPLICFC','X');
//
if not Bvente then
begin
	LaTOB.putValue('BLP_PHASETRA',THEdit(GetCOntrol('PHASE')).text);
end else
begin
  LaTOB.putValue('GL_COEFFG',Coeffg.value-1);

	if BLOQUETARIF.Checked then LaTOB.putValue('GL_BLOQUETARIF','X')
  											 else LaTOB.putValue('GL_BLOQUETARIF','-');
  LaTOB.PutValue('GL_COEFMARG',coefMarg.value);
  LaTOB.PutValue('POURCENTMARG',Arrondi((LaTOB.GetValue('GL_COEFMARG')-1)*100,2));

end;
TheTob:=LaTOB ;
end ;

procedure TOF_BTCOMPLIGNE.calculeCoefs;
var Montantmarg,PourcentMarg,MontantFrais : double;
begin
// Calcul
//if PA.value <> 0 then coeffg.Value := PR.value / PA.value else Coeffg.value := 0;
//
if Qte.value = 0 then
begin
  MontantFrais := Arrondi(MontantPR - MontantPA,V_PGI.okdecP);
end else
begin
  MontantFrais := Arrondi((MontantPR - MontantPA)/Qte.value,V_PGI.okdecP);
end;
SetControlProperty ('MONTANTFRAIS','value',MontantFrais);
//
if LaTOB.GetValue('GL_BLOQUETARIF')='X' then
begin
  if PR.value <> 0 then coefMarg.value := arrondi(PUHT.value / PR.value,4) else CoefMarg.value := 0;
end;
//
MontantMarg := PUHT.value - PR.value;
SetControlProperty ('MONTANTMARG','value',montantMarg);
//
if GetParamSocSecur('SO_BTGESTIONMARQ', False) then
begin
  if PUHT.value <> 0 then PourcentMarg := Arrondi((MontantMarg / PUHT.value)*100,4) else PourcentMarg := 0;
end else
begin
  if PR.value <> 0 then PourcentMarg := Arrondi((MontantMarg / PR.value)*100,4) else PourcentMarg := 0;
end;
SetControlProperty ('POURCENTMARG','value',PourcentMarg);
end;

procedure TOF_BTCOMPLIGNE.SetComponents;
var Indice : integer;
		UneTaxe : TOB;
begin
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    Unetaxe  := VH_GC.TOBParamTaxe.detail[Indice];
    // non paramétré
    THLabel (GetControl('TGL_FAMILLETAXE'+inttostr(Indice+1))).Visible := false;
    THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).Visible := false;
    THNumEdit(GetCOntrol('MILLIEME'+InttoStr(indice+1))).Visible := false;
    THLabel(GetCOntrol('MILLI'+InttoStr(indice+1))).Visible := false;
    if UneTaxe.getValue('BPT_TYPETAXE') <> '' then
    begin
      if UneTaxe.getValue('BPT_TYPETAXE') = 'TVA' Then
      begin
      //
      	if (Indice > 0) and ((TOBRepart = nil)  or (TOBRepart.detail.count = 0 )) then continue;
      //
      	THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).DataType := 'TTTVA';
      	THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).Visible := True;
      	THLabel (GetControl('TGL_FAMILLETAXE'+inttostr(Indice+1))).Caption := UneTaxe.getValue('BPT_LIBELLE');
      	THLabel (GetControl('TGL_FAMILLETAXE'+inttostr(Indice+1))).visible := true;
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
      	THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).DataType := 'TTTPF';
      	THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).Visible := True;
      	THLabel (GetControl('TGL_FAMILLETAXE'+inttostr(Indice+1))).Caption := UneTaxe.getValue('BPT_LIBELLE');
      	THLabel (GetControl('TGL_FAMILLETAXE'+inttostr(Indice+1))).visible := True;
      end;
      if (TOBRepart <> nil) and (TOBRepart.detail.count > 0) then
         THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).Enabled := false;
    end;
  end;
  APPLICCOEFFR.checked := not (LaTob.getValue('GLC_NONAPPLICFRAIS')='X');
  APPLICCOEFFC.checked := not (LaTob.getValue('GLC_NONAPPLICFC')='X');
  APPLICCOEFFG.checked := not (LaTob.getValue('GLC_NONAPPLICFG')='X');
end;

procedure TOF_BTCOMPLIGNE.OnLoad ;
Var Nature,TiersAff : string;
begin
  Inherited ;
  TOBPiece := LaTOB.Parent;
  if (TOBPiece.GetString('GP_NATUREPIECEG') = 'BCE') and (VH_GC.BTCODESPECIF = '001') then
  begin
    if (IsDebutParagraphe(LaTOB)) and (LaTOB.GetInteger('GL_NIVEAUIMBRIC')=1) then
    begin
      SetControlVisible('LGL_TYPECADENCE',true);
      SetControlVisible('GL_TYPECADENCE',true);
    end;
  end;
  ApplicFRST := (TOBPIece.GetValue('GP_APPLICFGST')='X');
  ApplicFCST := (TOBPIece.GetValue('GP_APPLICFCST')='X');
  PrestationST := IsPrestationST (LaTOb);
  if TOBPiece.getValue('GP_COEFFR')<> 0 then FRFixed := true else FRFixed := false;
  if TOBPiece.getValue('GP_COEFFC')<> 0 then FCFixed := true else FCFixed := false;

	if LaTOB.getValue('GL_PUHT') <> 0 then RATIOTTC.value := LaTOB.getValue('GL_PUTTC')/LaTOB.getValue('GL_PUHT')
  																	else RATIOTTC.value := 0;
  //
  BloquePV := (LaTOB.GetValue('GL_BLOQUETARIF')='X');
  DEV.Code:=laTob.getvalue('GL_DEVISE') ;
  GetInfosDevise(DEV) ;
  DEV.Taux:=LaTob.getValue('GL_TAUXDEV');

  TypeventilOuv := LaTOB.getValue('GL_TYPENOMENC');
  TypeArt := LaTOB.GetValue('GL_TYPEARTICLE');
  EnHt := (LaTOB.GetValue('GL_FACTUREHT')='X');
  AncAffaire := LaTOB.GetValue('GL_AFFAIRE');
if ctxAffaire in V_PGI.PGIContexte then
    BEGIN
    ChargeCleAffaire (Nil,THEDIT(GetControl('GL_AFFAIRE1')), THEDIT(GetControl('GL_AFFAIRE2')),
    THEDIT(GetControl('GL_AFFAIRE3')),THEDIT(GetControl('GL_AVENANT')),Nil, taCreat,GetControlText('GL_AFFAIRE'),False);
    END;
// VARIANTE
(* if copy(Latob.getValue('GL_TYPELIGNE'),1,2) = 'DP' Then *)
   if IsDebutParagraphe(LaTob) Then
   Begin
   Setcontrolvisible('PRIX',False);
   Setcontrolvisible('INFORMATION',False);
   Setcontrolvisible('GBBLOCNOTE',False);
   SetCOntrolVisible ('TGL_REFARTSAISIE',false);
   SetCOntrolVisible ('GL_REFARTSAISIE',false);
   SetCOntrolVisible ('TGL_REFARTTIERS',false);
   SetCOntrolVisible ('GL_REFARTTIERS',false);
   SetCOntrolVisible ('GL_LIBCOMPL',false);
   SetCOntrolVisible ('TGL_REFARTBARRE',false);
   SetCOntrolVisible ('GL_REFARTBARRE',false);
   SetCOntrolVisible ('G_COMM',false);
   end;
if LaTOB.GetValue('GL_IDENTIFIANTWOL') = -1 then TcheckBox(getControl('LIVRAISONFOUR')).checked := true
                                            else TcheckBox(getControl('LIVRAISONFOUR')).checked := false;

if LaTOB.getValue('GL_NATUREPIECEG')='ETU' then TcheckBox(getControl('LIVRAISONFOUR')).Visible  := false;
if ((TypeArt='ARP') and (Bvente)) or ((TypeArt='OUV') and (TypeVentilOuv = 'OUV')) then
begin
	SetControlText ('GB_INFOTAXE',TraduireMemoire('Dépendant du détail'));
  if (TOBrepart = nil) OR (TOBRepart.detail.count = 0) then
  begin
    SetControlText ('FAMILLETAXE1','');
    SetControlText ('FAMILLETAXE2','');
    SetControlText ('FAMILLETAXE3','');
    SetControlText ('FAMILLETAXE4','');
    SetControlText ('FAMILLETAXE5','');
  end else
  begin
    SetControlText ('FAMILLETAXE1',LaTOB.getValue('GL_FAMILLETAXE1'));
    SetControlText ('FAMILLETAXE2',LaTOB.getValue('GL_FAMILLETAXE2'));
    SetControlText ('FAMILLETAXE3',LaTOB.getValue('GL_FAMILLETAXE3'));
    SetControlText ('FAMILLETAXE4',LaTOB.getValue('GL_FAMILLETAXE4'));
    SetControlText ('FAMILLETAXE5',LaTOB.getValue('GL_FAMILLETAXE5'));
  end;
end else
begin
  SetControlText ('FAMILLETAXE1',LaTOB.getValue('GL_FAMILLETAXE1'));
  SetControlText ('FAMILLETAXE2',LaTOB.getValue('GL_FAMILLETAXE2'));
  SetControlText ('FAMILLETAXE3',LaTOB.getValue('GL_FAMILLETAXE3'));
  SetControlText ('FAMILLETAXE4',LaTOB.getValue('GL_FAMILLETAXE4'));
  SetControlText ('FAMILLETAXE5',LaTOB.getValue('GL_FAMILLETAXE5'));
	SetControlText ('GB_INFOTAXE','');
end;
if (LaTOB.getValue('GL_TYPELIGNE')<>'ART') and (copy(Latob.GetString('GL_TYPELIGNE'),1,2)<>'DP')  then
begin
  Setcontrolvisible('GL_DOMAINE',false);
  Setcontrolvisible('TGL_DOMAINE',false);
end;
if (LaTOB.getValue('GL_TYPEARTICLE') <> 'PRE')  then Setcontrolvisible('GB_RESSOURCE',false);
LaRessource := LaTOB.GetValue('GL_RESSOURCE');
Nature := LaTOB.GetValue('GL_NATUREPIECEG');
CodeTiers := LaTOB.GetValue('GL_TIERS');
if (Nature = VH_GC.AFNatAffaire ) or (Nature = VH_GC.AFNatProposition) then PieceAffaire := true
                                                                       else PieceAffaire := false;
if PieceAffaire then
   BEGIN
   SetControlEnabled ('GL_AFFAIRE1', False);
   SetControlEnabled ('GL_AFFAIRE2', False);
   SetControlEnabled ('GL_AFFAIRE3', False);
   setControlVisible ('BSELECTAFF1', False);
   setControlVisible ('BRAZAFFAIRE', False);
   END;
if (LaTOB.GetValue('GL_AFFAIRE') <> '') and Not(bVente) then
   begin
   TiersAff := GetChampsAffaire (LaTOB.GetValue('GL_AFFAIRE'), 'AFF_TIERS');
   SetControlText('TIERSAFFAIRE',tiersAff);
   end;
THValComboBox (GetControl('GL_DEPOT')).Enabled := true;
THValComboBox (GetControl('GL_DEPOT')).color := clWindow;
if EnHt then
begin
	PUTTCDEV.Enabled := false;
	PUTTCDEV.color := clBtnFace;
	PUHTDEV.Enabled := True;
	PUHTDEV.color := ClWindow;
end else
begin
	PUHTDEV.Enabled := false;
	PUHTDEV.color := clBtnFace;
	PUTTCDEV.Enabled := True;
	PUTTCDEV.color := ClWindow;
end;
if (Ouvrage) and (not isExistsArticle (trim(GetParamsoc('SO_BTECARTPMA')))) then
begin
	PUHTDEV.Enabled := false;
	PUHTDEV.color := clBtnFace;
	PUTTCDEV.Enabled := false;
	PUTTCDEV.color := clBtnFace;
end;
if BloquePv Then
begin
	PUHTDEV.Enabled := false;
	PUHTDEV.color := clBtnFace;
	PUTTCDEV.Enabled := false;
	PUTTCDEV.color := clBtnFace;
	COEFMARG.Enabled := false;
	COEFMARG.color := clBtnFace;
end;

if (ouvrage) then
begin
  COEFFG.enabled := false;
	COEFFG.color := clBtnFace;
  PA.enabled := false;
	PA.color := clBtnFace;
end;

if ((FRFixed) and (PrestationST) and (ApplicFRST)) then
begin
  PR.enabled := false;
	PR.color := clBtnFace;
end;

if ((FRFixed) and (not prestationST) and (APPLICCOEFFR.checked)) or (ouvrage) then
begin
  PR.enabled := false;
	PR.color := clBtnFace;
end;

if ((FCFixed) and (not prestationST) and (APPLICCOEFFC.checked)) or (ouvrage) then
begin
  PR.enabled := false;
	PR.color := clBtnFace;
end;

if ((FCFixed) or (FRFIxed) or (ouvrage)) then
begin
  PR.enabled := false;
	PR.color := clBtnFace;
end;

if (not APPLICCOEFFG.checked) or (Ouvrage) then
begin
  COEFFG.enabled := false;
	COEFFG.color := clBtnFace;
end;

if not Bvente then
begin
	if (not VarIsNull(LaTOB.GetValue ('BLP_PHASETRA'))) and (LaTOB.GetValue ('BLP_PHASETRA')<>'') then
  begin
    THEdit(GetCOntrol('PHASE')).text := LaTOB.GetValue ('BLP_PHASETRA');
  end;
end else
begin
	BLOQUETARIF.Checked := ( laTOB.getValue('GL_BLOQUETARIF')='X');
	if (not VarIsNull(LaTOB.GetValue ('GLC_FROMBORDEREAU'))) then
  begin
  	TcheckBox(GetControl('FROMBORDEREAU')).Checked := (LaTOB.GetValue ('GLC_FROMBORDEREAU')='X');
    if TcheckBox(GetControl('FROMBORDEREAU')).Checked then BLOQUETARIF.Enabled := false;
  end
end;

Getvalues;
// Ne faire qu'une fois pour recup les coefs
if Ouvrage then
begin
  if MontantPaFG <> 0 then CoefFG.value := Arrondi((MontantFG+MontantPAFG)/(MontantPAFG),4) else CoefFg.value := 0;
end else
begin
  CoefFG.value := LaTOB.GetValue('GL_COEFFG')+1;
end;
if (MontantPAFC <> 0) and (MontantFC <> 0) then CoeffC.value := Arrondi((MontantFC+MontantFG+MontantPAFC)/(MontantPAFC+MontantFG),4)
																					 else CoeffC.value := 0;
if (MontantPAFR <> 0) and (MontantFR <> 0) then Coeffr.value := Arrondi((MontantFG+MontantFR+MontantPAFR+MontantFC)/(MontantPAFG+MontantFC+MontantFG),4)
																					 else Coeffr.value := 0;
calculeCoefs;

if IsDejaFacture then
begin
	SetCOntrolEnabled('TVA',false);
	SetCOntrolEnabled('COEFMARG',false);
	SetCOntrolEnabled('GL_PUHTDEV',false);
end;

//uniquement en line
{*  TGRoupBox(GetControl('GSTOCK')).Visible := false;
  TGRoupBox(GetControl('G_COMM1')).Visible := false;
  TCheckBox(GetControl('LIVRAISONFOUR')).Visible := false;
  TCheckBox(GetControl('APPLICCOEFFG')).Visible := false;
  TCheckBox(GetControl('FROMBORDEREAU')).Visible := false;
  TTabSheet(GetControl('PRIX')).TabVisible := false;
*}

	if ThEdit(GetControl('GL_AFFAIRE1')).Enabled then ThEdit(GetControl('GL_AFFAIRE1')).SetFocus
  else if ThEdit(GetControl('GL_AFFAIRE2')).Enabled then ThEdit(GetControl('GL_AFFAIRE2')).SetFocus
  else if ThEdit(GetControl('GL_AFFAIRE3')).Enabled then ThEdit(GetControl('GL_AFFAIRE3')).SetFocus;
  if isOuvrage(LaTOB) and (LaTOB.Getboolean('GL_GESTIONPTC')) then
  begin
    ThEdit(GetControl('GL_PUHTDEV')).Enabled := false;
    ThEdit(GetControl('COEFMARG')).Enabled := false;
  end;

end ;

procedure TOF_BTCOMPLIGNE.SetDecimals;
begin

	THNumEdit(GetControl('GL_QTEFACT')).NumericType := ntDecimal;
	THNumEdit(GetControl('GL_QTEFACT')).Decimals := V_PGI.OkdecQ;
  //
	THNumEdit(GetControl('GL_DPA')).NumericType := ntDecimal;
	THNumEdit(GetControl('GL_DPA')).Decimals := V_PGI.OkdecP;
  //
	THNumEdit(GetControl('GL_DPR')).NumericType := ntDecimal;
	THNumEdit(GetControl('GL_DPR')).Decimals := V_PGI.OkdecP;
  //
	THNumEdit(GetControl('MONTANTFRAIS')).NumericType := ntDecimal;
	THNumEdit(GetControl('MONTANTFRAIS')).Decimals := V_PGI.OkdecP;
  //
	THNumEdit(GetControl('GL_PUHTBASE')).NumericType := ntDecimal;
	THNumEdit(GetControl('GL_PUHTBASE')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('MONTANTMARG')).NumericType := ntDecimal;
	THNumEdit(GetControl('MONTANTMARG')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('GL_PUHTDEV')).NumericType := ntDecimal;
	THNumEdit(GetControl('GL_PUHTDEV')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('GL_PUTTCDEV')).NumericType := ntDecimal;
	THNumEdit(GetControl('GL_PUTTCDEV')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('GL_PUHT')).NumericType := ntDecimal;
	THNumEdit(GetControl('GL_PUHT')).Decimals := V_PGI.OkdecP;
	THNumEdit(GetControl('GL_PUTTC')).NumericType := ntDecimal;
	THNumEdit(GetControl('GL_PUTTC')).Decimals := V_PGI.OkdecP;
end;

procedure TOF_BTCOMPLIGNE.OnArgument (S : String ) ;
begin
  Inherited ;
  AppliqueFontDefaut (THRichEditOle(GetControl('GL_BLOCNOTE')));
  TOBrepart := nil;

  if LaTOB.data <> nil then TOBRepart := TOB(laTOB.data);
  PuDEVModifie := false;
	Ouvrage := ( ((LaTOB.GetValue('GL_TYPEARTICLE')='OUV') OR (LaTOB.GetValue('GL_TYPEARTICLE')='ARP')) and
  						(LaTOB.GetValue('GL_INDICENOMEN')>0));
  IsDejaFacture := false;
  if ((LaTOB.GetValue('GL_QTEPREVAVANC') <> 0) or (LaTOB.GetValue('GL_QTESIT') <> 0)) and
  		(LaTOB.GetValue('GL_NATUREPIECEG')='DBT')  then IsDejaFacture := true;

  bChangetiers := Not(bVente);
  SetcontrolVisible('TIERSAFFAIRE', Not(bVente));
  SetcontrolVisible('LIBTIERSAFFAIRE', Not(bVente));
  SetcontrolVisible('BEFFACEAFF1', Not(bVente));
  SetControlVisible ('INFORMATION',False);
  TTAbSheet(getControl('PRIX')).tabvisible := Bvente;
  if Not(bVente) then SetControlText('TGL_TIERS','&Fournisseur');
  GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GL_LIBREART', 10, '');
  Qte := THNumEdit(GetCOntrol('GL_QTEFACT'));
  PuHTDEV := THNumEdit(GetControl ('GL_PUHTDEV')); PuHtDev.OnExit := PuHTDevExit;
  PuHT := THNumEdit(GetControl ('GL_PUHT'));
  PA := THNumEdit(GetControl ('GL_DPA'));
  PR := THNumEdit(GetControl ('GL_DPR'));
  PUTTC := THNumEdit(GetControl ('GL_PUTTC'));
  PUTTCDEV := THNumEdit(GetControl ('GL_PUTTCDEV')); PuTTCDev.onExit := PUTTCDevExit;
  RATIOTTC := THNumEdit(GetControl ('RATIOTTC'));
  COEFFR := THNumEdit(GetControl ('COEFFR'));
  COEFFC := THNumEdit(GetControl ('COEFFC'));
  COEFMARG := THNumEdit(GetControl ('COEFMARG'));
  coefMarg.value := LaTOB.GetValue('GL_COEFMARG');
  COEFMARG.OnExit := COEFMARGExit;
	THEdit(GetCOntrol('PHASE')).text := '';

  APPLICCOEFFR := TCheckBox(GetCOntrol('APPLICCOEFFR'));
  APPLICCOEFFC := TCheckBox(GetCOntrol('APPLICCOEFFC'));
  APPLICCOEFFG := TCheckBox(GetCOntrol('APPLICCOEFFG'));
  COEFFG := THNumEdit(GetControl('COEFFG'));
  if not Ouvrage then
  begin
//    COEFFC.OnExit := COEFFCExit;
    COEFFG.Onexit := COEFFGExit;
    COEFFR.OnExit := COEFFRExit;
    PA.OnExit := PAExit;
    PR.OnExit := PRExit;
  end;
  if bvente then
  begin
  	BLOQUETARIF := TCheckBox(GetControl('BLOQUETARIF'));
  	BLOQUETARIF.visible := True;
    BLOQUETARIF.onclick := BloqueTarifClick;
  	TToolbarButton97(GetControl('BSELECTAFF1')).visible := false;
    TToolBarButton97(getControl('BRAZAFFAIRE')).Visible := False;
  	TCheckBox(GetControl('FROMBORDEREAU')).visible := True;
    	THEdit(GetCOntrol('GP_VENTEACHAT')).text:='VEN';
  end else
  begin
  	THLabel(getControl('TPHASES')).Visible := true;
  	TToolBarButton97(getControl('BRECHPHASE')).Visible := true;
    TToolBarButton97(getControl('BRECHPHASE')).OnClick := BrechPhaseClick;
    THedit(GetControl('PHASE')).OnChange := PhaseChange;
  	THLabel(getControl('LIBPHASE')).Visible := true;
    	THEdit(GetCOntrol('GP_VENTEACHAT')).text:='ACH';
  end;
  //
  TToolBarButton97(getControl('BRAZAFFAIRE')).OnClick := BRazAffaireClick;

  //
  SetDecimals;
	SetComponents;

  APPLICCOEFFR.OnClick := APPLICCOEFFRClick;
  APPLICCOEFFC.OnClick := APPLICCOEFFCClick;
  APPLICCOEFFG.OnClick := APPLICCOEFFGClick;

  if Pos('CONSULTPRIX',S) > 0 then ModeConsultPrix;

	THEdit(GetCOntrol('GL_AFFAIRE0')).text := Copy(LaTOB.GetValue('GL_AFFAIRE'),1,1);
  TCheckBox(getControl('GL_GESTIONPTC')).OnClick := GESTIONPTCClick;

end ;

procedure TOF_BTCOMPLIGNE.ModeConsultPrix;
begin
	TTabSheet (GetControl('PRIX')).Enabled := false;
end;

procedure TOF_BTCOMPLIGNE.PAExit (Sender : TObject);
begin
  MontantPA := arrondi(PA.Value * QTe.Value,V_PGI.okDECV);
  if not Ouvrage then
  begin
  	if APPLICCOEFFG.Checked then MontantPAFG := montantPA else MontantPAFG := 0;
    if APPLICCOEFFC.Checked then MontantPAFC := MontantPA else MontantPAFC := 0;
    if APPLICCOEFFR.Checked then MontantPAFR := MontantPA else MontantPAFR := 0;
  end;
  CalcPrLocal;
  CalculeCoefs;
end;

procedure TOF_BTCOMPLIGNE.PRExit (Sender : TObject);
begin
	if PA.value <> 0 then COEFFG.Value := (PR.Value / PA.value);
  CalcPrLocal;
  CalculeCoefs;
end;

procedure TOF_BTCOMPLIGNE.COEFFRExit (Sender : TObject);
begin
  CalcPrLocal;
  calculeCoefs;
end;


procedure TOF_BTCOMPLIGNE.COEFMARGExit (Sender : TObject);
begin
  CalcPrLocal;
  CalculeCoefs;
end;

procedure TOF_BTCOMPLIGNE.PuHtDevExit (Sender : TObject);
var savPb : boolean;
begin
  DeviseToPivot ( PUHT,PUHtDEv) ;
  SavPb := BloquePV;
  BloquePV := true;
  CalcPrLocal;
  BloquePv := SavPb;
  CalculeCoefs;
  PuDEVModifie := true;
  //BLOQUETARIF.Enabled := false;  mis en commentaire par brl 25102010 car aucune raison apparente
end;

procedure TOF_BTCOMPLIGNE.OnClose ;
var EtatAffaire : String;
    Aff0,Aff,Aff1,Aff2,Aff3,Aff4,Tiers    : THEdit;
    Part0,Part1,Part2,Part3,part4,Domaine : string;
begin
  Inherited ;

  //FV1 : 18/09/2013 - FS#665 - En saisie de pièce achats, le contrôle sur affaire acceptée n'est pas fait dans le complément ligne
  if (GetParamSocSecur('SO_BTINTERDIREACHATS',False)) Then
  begin
    if GetControlText('GP_VENTEACHAT')= 'ACH' then
    begin
      //FV1 : 10/06/2014 - FS#921 - DELABOUDINIERE : Revoir les contrôles sur appels et contrats en fonction du code état
      if not ControleAffaire(THEdit(GetControl('GL_AFFAIRE')).Text, Ecran.caption,'ACH') then
      begin
        BTPCodeAffaireDecoupe(AncAffaire,Part0, Part1,Part2,Part3,Part4,TaModif,False);
        THEdit(GetControl('GL_AFFAIRE0')).Text  := Part0;
        THEdit(GetControl('GL_AFFAIRE1')).Text  := Part1;
        THEdit(GetControl('GL_AFFAIRE2')).Text  := Part2;
        THEdit(GetControl('GL_AFFAIRE3')).Text  := Part3;
        THEdit(GetControl('GL_AVENANT')).Text   := Part4;
      end;
    end;
  end;

end ;

Procedure AGLAffecteRessourceLigne ( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
BEGIN
F:=TForm(Longint(Parms[0]));
if (F is TFVierge) then LaTof:=TFVierge(F).Latof else laTof := nil;
if (LaTof is TOF_BTCOMPLIGNE) then TOF_BTCOMPLIGNE(LaTof).AffecteRessourceLigne;
END;

procedure TOF_BTCOMPLIGNE.ValideChamps;
begin
NextControl (Ecran,true);
end;

procedure TOF_BTCOMPLIGNE.DeviseToPivot (ZOnePivot,ZoneDev : THnumEdit);
begin

  if DEV.Code <> V_PGI.DevisePivot  then
  begin
      ZOnePivot.Value := DEVISETOEURO (ZoneDev.Value,DEV.Taux,DEV.quotite);
  end else
  begin
      ZOnePivot.Value := ZoneDev.Value;
  end;

end;

procedure TOF_BTCOMPLIGNE.PuTTCDevExit(Sender: TObject);
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
  PuDEVModifie := true;
  //BLOQUETARIF.Enabled := false;  mis en commentaire par brl 25102010 car aucune raison apparente
end;

procedure TOF_BTCOMPLIGNE.bSelectAff1Click(Sender: TObject);
begin
  if TOBPiece.getValue ('GP_VENTEACHAT')='ACH' then  bchangeStatut := True;
  inherited;
  if Existeaffaire (THEDIT(GetControl('GL_AFFAIRE')).Text ,'') then
  begin
  	if GetParamSoc ('SO_BTLIVCHANTIER') then
    begin
      if (LaTOB.GetString('GL_NATUREPIECEG')<>'CBT') or
         ((LaTOB.GetString('GL_NATUREPIECEG')='CBT') and (not GetParamSocSecur('SO_BTLIVBESOINDEF',false))) then
      begin
        TcheckBox(getControl('LIVRAISONFOUR')).checked := true;
      end;
    end;
  end;
end;

procedure TOF_BTCOMPLIGNE.BrechPhaseClick(Sender: Tobject);
var Code : string;
begin
  if LaTOB.getValue('GL_AFFAIRE') = '' then exit;

  if SelectionPhase (LaTOB.getValue('GL_AFFAIRE'),Code) then
  begin
  	 THEdit(GetControl('PHASE')).Text := Code;

  end;
end;

procedure TOF_BTCOMPLIGNE.PhaseChange(Sender: TObject);
begin
	THLabel(getControl('LIBPHASE')).Caption := GetLibellePhase (LaTOB.getValue('GL_AFFAIRE'),THEdit(GetControl('PHASE')).Text);
end;

procedure TOF_BTCOMPLIGNE.BloqueTarifClick (Sender : Tobject);
var Autorise : boolean;
begin
	Autorise := not BLOQUETARIF.checked ;
	PUHTDEV.Enabled := Autorise;
	PUTTCDEV.Enabled := Autorise;
	COEFMARG.Enabled := Autorise;
  if Autorise then
  begin
  	PUHTDEV.color := ClWindow;
    PUTTCDEV.color := ClWindow;
    COEFMARG.color := ClWindow;
    BloquePV := false;
  end else
  begin
  	PUHTDEV.color := clBtnFace;
    PUTTCDEV.color := clBtnFace;
    COEFMARG.color := clBtnFace;
    BloquePV := true;
  end;
end;


procedure TOF_BTCOMPLIGNE.CalcPrLocal;
begin
  if APPLICCOEFFG.checked then
  BEGIN
  	if COEFFG.Value  <> 0 then MontantFG := arrondi((MontantPAFG * COEFFG.Value) - MontantPAFG,V_PGI.OkdecP)
    											else MontantFG := 0;
  END else
  BEGIN
    MontantFG := 0;
  END;

	if APPLICCOEFFC.checked then
  begin
  	if COEFFC.Value  <> 0 then MontantFC := arrondi(((MontantPAFC+MontantFG) * COEFFC.Value) - (MontantPAFC+MontantFG),V_PGI.okdecP)
    											else MontantFC := 0;
  end else
  begin
  	MontantFC := 0;
  end;

	if APPLICCOEFFR.checked then
  begin
  	if COEFFR.Value <> 0 then montantFR := arrondi(((MontantPAFR + MontantFC + MontantFG) * COEFFR.Value) - (MontantPAFR + MontantFC+ MontantFG),V_PGI.OkdecP)
    									   else MontantFR := 0;
  end else
  begin
  	montantFR := 0;
  end;

  if MontantFG+MontantFR+MontantFC <> 0 then
  begin
  	MontantPR := MontantPA+MontantFG+MontantFC+MontantFR;
    if Qte.value = 0 then
    begin
      PR.Value := 0;
    end else
    begin
      PR.Value := arrondi(MontantPR / Qte.value,V_PGI.OKDecp);
    end;
  end else
  begin
  	MontantPR := MontantPA;
    if Qte.value = 0 then
    begin
      PR.Value := 0;
    end else
    begin
      PR.Value := arrondi(MontantPR / Qte.value,V_PGI.OKDecp);
    end;
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
  	if Pr.Value <> 0 then COEFMARG.Value := arrondi (PUHt.value/Pr.value,4) else COEFMARG.value := 0;
  end;
end;


procedure TOF_BTCOMPLIGNE.Getvalues;
begin
	MontantPA := LaTOB.GetValue('GL_MONTANTPA');
	MontantPAFG := LaTOB.GetValue('GL_MONTANTPAFG');
  MontantFG := LaTOB.GetValue('GL_MONTANTFG');
	MontantPAFC := LaTOB.GetValue('GL_MONTANTPAFC');
  MontantFC := LaTOB.GetValue('GL_MONTANTFC');
  MontantPAFR := LaTOB.GetValue('GL_MONTANTPAFR');
  montantFR := LaTOB.GetValue('GL_MONTANTFR');
  MontantPR := LaTOB.GetValue('GL_MONTANTPR')
end;


procedure TOF_BTCOMPLIGNE.APPLICCOEFFRClick(Sender: Tobject);
begin
  if not APPLICCOEFFG.checked then
  begin
  	MontantFR := 0;
  end else
  begin
  	if TOBPiece.getValue('GP_COEFFR') <> 0 then
    begin
    	if Ouvrage then
      begin
      	MontantPAFR := LaTOB.GetValue('GL_MONTANTPAFR');
      end else
      begin
      	MontantPAFR := LaTOB.GetValue('GL_MONTANTPA');
      end;
    	if ((PrestationST) and (ApplicFRSt)) or (not PrestationST) then
      begin
        MontantFR := Arrondi((MontantPAFR + MontantFC + MontantFG) * TOBPiece.getValue('GP_COEFFR'),V_PGI.OkdecP);
        if MontantPAFR <> 0 then
        begin
          COEFFR.Value := Arrondi((MontantPAFR+MontantFC+MontantFR+MontantFG) / (MontantPAFR+MontantFc+MontantFG),4);
        end;
      end;
    end;
  end;
  CalcPrLocal;
  calculeCoefs;
end;

procedure TOF_BTCOMPLIGNE.APPLICCOEFFCClick(Sender: Tobject);
begin
  if not APPLICCOEFFC.checked then
  begin
  	MontantFC := 0;
  end else
  begin
  	if TOBPiece.getValue('GP_COEFFC') <> 0 then
    begin
    	if Ouvrage Then
      begin
    		MontantPAFC := LaTOB.GetValue('GL_MONTANTPAFC');
      end else
      begin
    		MontantPAFC := LaTOB.GetValue('GL_MONTANTPA');
      end;
    	if ((PrestationST) and (ApplicFcSt)) or (not PrestationST) then
      begin
        MontantFC := Arrondi((MontantPAFC+MontantFG) * TOBPiece.getValue('GP_COEFFC'),V_PGI.okdecP);
        if MontantPAFC <> 0 then
        begin
          COEFFC.Value := arrondi((MontantPAFC+MontantFC+MontantFG) / (MontantPAFC+MontantFG),4);
        end;
      end else
      begin
      end;
    end;
  end;
  CalcPrLocal;
  calculeCoefs;
end;      

{*procedure TOF_BTCOMPLIGNE.Setvalues;
begin
  if not Ouvrage then
  begin
		LaTOB.PutValue('GL_MONTANTPA',MontantPA);
    LaTOB.PutValue('GL_MONTANTPAFG',MontantPaFG);
    LaTOB.PutValue('GL_MONTANTPAFC',MontantPaFC);
    LaTOB.PutValue('GL_MONTANTPAFR',MontantPaFR);
  end;
  LaTOB.PutValue('GL_MONTANTFG',MontantFG);
  LaTOB.PutValue('GL_MONTANTFC',MontantFC);
  LaTOB.PutValue('GL_MONTANTFR',MontantFR);
  if (MontantPAFC+MontantFG) <> 0 then LaTOB.PutValue('GL_COEFFC',Arrondi(MontantFC/(MontantPAFC+MontantFG),4));
  if (MontantPAFR+MontantFC+MontantFG) <> 0 then LaTOB.PutValue('GL_COEFFR',Arrondi(MontantFR/(MontantPAFR+MontantFC+MontantFG),4));
  LaTOB.PutValue('GL_MONTANTPR',MontantPR);
end;*}

procedure TOF_BTCOMPLIGNE.APPLICCOEFFGClick(Sender: Tobject);
begin
  if not APPLICCOEFFG.checked then
  begin
  	MontantFG := 0;
  end else
  begin
  	if LaTOB.getValue('GL_COEFFG') <> 0 then
    begin
    	if Ouvrage Then
      begin
    		MontantPAFG := LaTOB.GetValue('GL_MONTANTPAFG');
      end else
      begin
    		MontantPAFG := LaTOB.GetValue('GL_MONTANTPA');
      end;
      MontantFG := Arrondi(MontantPAFG* LaTOB.getValue('GL_COEFFG'),V_PGI.okdecP);
      if MontantPaFG <> 0 then CoeffG.value := Arrondi((MontantPaFG+MontantFG)/MontantPaFG,4);
    end;
  end;
  CalcPrLocal;
  calculeCoefs;
end;

procedure TOF_BTCOMPLIGNE.COEFFGExit(SENDER: TOBject);
begin
  if COEFFG.value <> 0 then
  begin
    if Ouvrage Then
    begin
      MontantPAFG := LaTOB.GetValue('GL_MONTANTPAFG');
    end else
    begin
      MontantPAFG := LaTOB.GetValue('GL_MONTANTPA');
    end;
    MontantFG := Arrondi((MontantPAFG* COEFFG.value)-MontantPAFg,V_PGI.okdecP);
    if MontantPaFG <> 0 then CoeffG.value := Arrondi((MontantPaFG+MontantFG)/MontantPaFG,4);
  end;
  CalcPrLocal;
  calculeCoefs;
end;

procedure TOF_BTCOMPLIGNE.BRazAffaireClick(Sender: TObject);
Var Aff0, Aff, Aff1, Aff2, Aff3, Aff4, Tiers : THEdit;
begin

  Aff:=THEdit(GetControl('GL_AFFAIRE'));
  Aff0:=THEdit(GetControl('GL_AFFAIRE0'));
  Aff1:=THEdit(GetControl('GL_AFFAIRE1'));
  Aff2:=THEdit(GetControl('GL_AFFAIRE2')); Aff3:=THEdit(GetControl('GL_AFFAIRE3'));
  Aff4:=THEdit(GetControl('GL_AVENANT'));

  Aff.Text  := '';
  Aff0.text := '';
  Aff1.text := '';
  Aff2.text := '';
  Aff3.text := '';
  Aff4.text := '';
  
end;

procedure TOF_BTCOMPLIGNE.GESTIONPTCClick(Sender: TObject);
begin
  if TCheckBox(geTControl('GL_GESTIONPTC')).Checked then
  begin
    if not isExistsArticle(trim(GetParamsoc('SO_BTECARTPMA'))) then
    begin
      PgiBox(TraduireMemoire('L''article d''écart est invalide ou non renseigné#13#10Veuillez le définir'), Traduirememoire('Gestion d''écart'));
    end;
    ThEdit(GeTControl('GL_PUHTDEV')).enabled := false;
    ThEdit(GetControl('COEFMARG')).Enabled := false;
  end else
  begin
    if not IsDejaFacture then
    begin
      ThEdit(GeTControl('GL_PUHTDEV')).enabled := true;
      ThEdit(GetControl('COEFMARG')).Enabled := true;
    end;
  end;
end;

Initialization
  RegisterAglProc( 'ChangeRessourceLigne', True ,1, AGLAffecteRessourceLigne);
  registerclasses ( [ TOF_BTCOMPLIGNE ] ) ;
end.
