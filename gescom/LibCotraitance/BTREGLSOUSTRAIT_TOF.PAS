{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/03/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREGLSOUSTRAIT ()
Mots clefs ... : TOF;BTREGLSOUSTRAIT
*****************************************************************}
Unit BTREGLSOUSTRAIT_TOF ;

Interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HSysMenu
  , HCtrls, HEnt1, HMsgBox, UTOF, vierge, UTOB, AglInit, LookUp, EntGC, SaisUtil, graphics
  , grids, windows, utilPGI, ent1, FactRG, ParamSoc
  {$IFDEF EAGLCLIENT}
  , MaineAGL
  {$ELSE}
  , db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main
  {$ENDIF}
  , M3FP, Messages, FactCpta,TiersUtil,
{$IFDEF BTP}
  FactAcompte,
{$ENDIF}
	utilcb,
	HPanel,
  HTB97;

Type
  TOF_BTREGLSOUSTRAIT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Erreur, OkFermer: boolean;
    PassCompta : boolean;
    FIndexFille: integer;
    TobPiece, TobTiers, TobAcomptes: Tob;
    TOBPieceTrait : TOB;
    Action: TActionFiche;
    IsReglement: boolean;
    // MODIF BTP
    TOBPieceRG: TOB;

    procedure BdeleteClick (Sender : Tobject);
    procedure ChangeDateAcc (Sender : Tobject);
    function GetLibelleRegl: string;
    procedure ChangeTypeSaisie (Sender : TObject);
    function FindPaiementSousTrait : Double ;
  end ;

Implementation
uses facttiers;

procedure TOF_BTREGLSOUSTRAIT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLSOUSTRAIT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLSOUSTRAIT.OnUpdate ;
var X: T_GCAcompte;
  QQ: TQuery;
  TOBClient,TOBFournisseur : TOB;

  procedure GereError(Num: integer; Msg: string);
  begin
    LastError := Num;
    LastErrorMsg := TraduireMemoire(Msg);
    TToolBarButton97(GetControl('BVALIDER')).Enabled := True;
  end;

begin
  inherited;
  TOBClient := TOB.Create('TIERS',nil,-1);
  TOBClient.Dupliquer(TobTiers,false,true);
  TOBFournisseur := TOB.Create('TIERS',nil,-1);
  RemplirTOBTiers (TOBFournisseur,GetControlText('GAC_FOURNISSEUR'),'',False);
  //
  TToolBarButton97(GetControl('BVALIDER')).Enabled := False;
  OkFermer := True;
  if Action = taConsult then Exit;
  LastErrorMsg := '';
  LastError := 0;
  Erreur := False;

  if GetControlText('GAC_LIBELLE') = '' then GereError(1, 'Le libellé est obligatoire')
  else if GetControlText('GAC_MODEPAIE') = '' then GereError(2, 'Le mode de paiement est obligatoire')
  else if (GetControlText('GAC_JALECR') = '') and (PassCompta) then GereError(3, 'Le journal est obligatoire')
  else if Valeur(GetControlText('GAC_MONTANTDEV')) <= 0 then GereError(4, 'Le montant doit être supérieur à zéro');
  // JT 09/10/2003, si CB, test si on doit faire le contrôle du n° et date
  if not PassCompta then
  begin
    if Valeur(GetControlText('GAC_MONTANTDEV')) > TOBPiece.GetValue('GP_TOTALTTCDEV') then
    begin
			GereError(8, 'Le montant excède celui du document');
		end;
  end;

  Erreur := (LastError <> 0);
  if Erreur then Exit;
  FillChar(X, Sizeof(X), #0);
  if PassCompta then
  begin
    X.JalRegle := getcontroltext('GAC_JALECR');
    X.CpteRegle := '';
    QQ := OpenSql('Select J_MODESAISIE from JOURNAL where J_JOURNAL="' + X.JalRegle + '" AND J_NATUREJAL="OD"', TRUE);
    if not QQ.EOF then
    begin
      if QQ.findField('J_MODESAISIE').asString <> '-' then
      begin
        LastError := 7;
        LastErrorMsg := 'Le journal doit être de type OD en mode pièce';
      end;
//      X.CpteRegle := QQ.fields[0].asstring;
    end;
  	ferme(QQ);
  end;
  if LastError <> 0 then
  begin
  	erreur := (LastError <> 0);
		ecran.ModalResult := MrCancel;
  	TToolBarButton97(GetControl('BVALIDER')).Enabled := true;
    exit;
  end;
  X.Libelle := getcontroltext('GAC_LIBELLE');
  X.ModePaie := getcontroltext('GAC_MODEPAIE');
  X.Montant := valeur(getcontroltext('GAC_MONTANTDEV'));
  X.IsReglement := true;
  X.IsContrepartie := false;
  X.CpteContre := '';
  X.LibelleContre := '';
  X.DateECr := getcontroltext('GAC_DATEECR');
  X.DateEche := getcontroltext('GAC_DATEECHEANCE');
  // Modif JT le 09/10/03 -- Gestion du tiers Facturé
//  X.AuxiFacture := AuxiFacture;
  // Modif BTP
  X.IsComptabilisable := PassCompta;

  X.IsModif := (Action = TAModif);
  if FIndexFille >= 0 then // cas d'une modif
  begin
    X.NumEcr := TobAcomptes.detail[FIndexFille].GetValue('GAC_NUMECR');
    X.LaTobAcc := TobAcomptes.detail[FIndexFille];
  end else
  begin
    X.numecr := 0;
    X.LaTobAcc := nil;
  end;
  // --
  TheTob := EnregistreReglSSTrait(TOBPiece, TOBClient,TOBFournisseur, X);
  if LaTOB.FieldExists('VALIDSOLO') then LaTOB.SetString('VALIDSOLO','X');
  //
  TOBClient.free;
  TOBFournisseur.free;
  //
end ;

procedure TOF_BTREGLSOUSTRAIT.OnLoad ;
var RG: THRadioGroup;
begin
  Inherited ;
  Erreur := False;
  OkFermer := False;
  if (Action = taConsult) or (Action = TaModif) then
  begin
    OkFermer := True;
    if FIndexFille < 0 then Exit;
    TobAcomptes.detail[FIndexFille].PutEcran(Ecran);
    exit;
  end;
  if Action = taCreat then
  begin
    RG := THRadioGroup(GetControl('GAC_ISREGLEMENT'));
    if IsReglement then RG.Value := 'X' else RG.Value := '-';
  end;
end ;

procedure TOF_BTREGLSOUSTRAIT.OnArgument (S : String ) ;
var St, Critere: string;
  x: integer;
  ChampMul, ValMul: string;
  // Modif BTP
  ComptaACCReglements : boolean;
  PieceComptaAccRegl : boolean;
  LibReglement : string;
  // --
begin
  TobTiers := LaTob.Detail[0];
  TobPiece := LaTob.Detail[0].Detail[0];
  TobAcomptes := LaTob.Detail[0].Detail[0].Detail[0];
  // Modif BTP
  TOBPieceRG := TOB(LaTob.data);
  TOBPieceTrait := TOB(TOBPieceRg.data);
  //
  PassCompta := true;

  ComptaACCReglements := GetParamSocSecur ('SO_BTCOMPTAREGL',false);
  PieceComptaAccRegl := (GetInfoParPiece (TOBpiece.getValue('GP_NATUREPIECEG'),'GPP_TYPEPASSACC')<>'AUC');
  if (not ComptaACCReglements) or (not PieceComptaAccRegl) then
  begin
  	PassCompta := false;
  end;
  St := S;
  Action := taModif;
  IsReglement := false;
  FIndexFille := (-1);
  repeat
    Critere := uppercase(ReadTokenSt(St));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'ACTION' then
        begin
          if ValMul = 'CREATION' then
          begin
            Action := taCreat;
          end;
          if ValMul = 'MODIFICATION' then
          begin
            Action := taModif;
          end;
          if ValMul = 'CONSULTATION' then
          begin
            Action := taConsult;
          end;
        end
        else if ChampMul = 'LATOB' then FIndexFille := strToInt(ValMul)
        else if ChampMul = 'FOURNISSEUR' then
        begin
          IsReglement := false;
          SetControltext('GAC_FOURNISSEUR', ValMul);
        end
        else if ChampMul = 'TITRE' then
        begin
          Ecran.caption := ValMul;
          UpdateCaption(Ecran);
          Ecran.refresh;
        end
          ;
      end;

    end;
  until Critere = '';
  IsReglement := true;
  SetControlEnabled('GAC_ISREGLEMENT', false);
  SetControlChecked ('GAC_ISREGLEMENT', true);
  THValComboBox (GetControl('GAC_MODEPAIE')).Value := GetParamSocSecur('SO_BTPAIESTDIRECTE','');
  if Action = taConsult then
  begin
    FicheReadOnly(Ecran);
  end else
  begin
    if Action = TaCreat then
    begin
      SetControlText ('GAC_DATEECR',DateToStr(Now));
      if not IsReglement then
      begin
        LibReglement := Copy('Acompte ' + DateToStr(TOBPiece.GetValue('GP_DATEPIECE')) + ' Pce. '+ IntToStr(TOBPiece.GetValue('GP_NUMERO')), 1, 35);
      end else
      begin
        LibReglement := Copy('Règlement ' + DateToStr(TOBPiece.GetValue('GP_DATEPIECE')) + ' Pce. '+ IntToStr(TOBPiece.GetValue('GP_NUMERO')), 1, 35);
      end;
      SetControlText('GAC_LIBELLE', LibReglement );
      ThEdit(GetControl('GAC_DATEECR')).OnChange := ChangeDateAcc;
    end;
    if Action = TaModif then
    begin // On autorise que la modif du montant
      THRadioGRoup(GetControl('GAC_ISREGLEMENT')).Enabled := false;
      THEdit(GetControl('GAC_MODEPAIE')).Enabled := false;
      THEdit(GetControl('GAC_DATEECR')).Enabled := false;
      THEdit(GetControl('GAC_JALECR')).Enabled := false;
      SetFocusControl('GAC_MONTANTDEV');
    end else SetFocusControl('GAC_MODEPAIE');
  end;
  if not PassCompta then
  begin
  	THEdit(GetControl('GAC_JALECR')).Visible := false;
    THLabel(GetControl('TJALECR')).Visible := false;
    if (Action <> taConsult) then
    begin
      TToolBarButton97(GetControl('BDELETE')).Visible := true;
      TToolBarButton97(GetControl('BDELETE')).onclick := BdeleteClick;
    end;
  end;
  THNumEdit(GetControl('GP_TOTALTTCDEV')).value := FindPaiementSousTrait;
  // --
	THRadioGroup(GetControl('GAC_ISREGLEMENT')).OnClick := ChangeTypeSaisie;

end ;

procedure TOF_BTREGLSOUSTRAIT.OnClose ;
begin
  Inherited ;
  if (not OkFermer) and (valeur(getcontroltext('GAC_MONTANTDEV')) <> 0) then
  begin
    if PGIAsk('Confirmez-vous l''abandon ?', Ecran.Caption) = mrNo then
    begin
      LastError := (-1);
      LastErrorMsg := '';
      Exit;
    end else
    begin
      LastError := 0;
      Exit;
    end;
  end;

  if Erreur then
  begin
    LastError := (-1);
    LastErrorMsg := '';
    Erreur := false;
    OkFermer := False;
  end;
end ;

procedure TOF_BTREGLSOUSTRAIT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLSOUSTRAIT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLSOUSTRAIT.BdeleteClick(Sender: Tobject);
begin
	if PGIAsk('Confirmez-vous la suppression ?', Ecran.Caption) = mryes then
  begin
    TobAcomptes.detail[FIndexFille].free;
    ecran.ModalResult:=mrCancel;
  end;
end;

procedure TOF_BTREGLSOUSTRAIT.ChangeDateAcc(Sender: Tobject);
begin
	SetControlText('GAC_LIBELLE',GetLibelleRegl);
end;

procedure TOF_BTREGLSOUSTRAIT.ChangeTypeSaisie(Sender: TObject);
begin
	SetControlText('GAC_LIBELLE',GetLibelleRegl);
end;


function TOF_BTREGLSOUSTRAIT.GetLibelleRegl: string;
begin
	result := Copy('Paiement ' + GetControlText ('GAC_DATEECR') + ' -ST- '+GetControlText('GAC_FOURNISSEUR'), 1, 35);
end;

function TOF_BTREGLSOUSTRAIT.FindPaiementSousTrait: Double;
var TOBPT : TOB;
begin
	Result := 0;
  TOBPT := TOBPieceTrait.findFirst(['BPE_FOURNISSEUR'],[GetControlText('GAC_FOURNISSEUR')],True);
  if TOBPT <> NIL then Result := TOBPT.GetDouble('BPE_MONTANTREGL');
end;

Initialization
  registerclasses ( [ TOF_BTREGLSOUSTRAIT ] ) ; 
end.
