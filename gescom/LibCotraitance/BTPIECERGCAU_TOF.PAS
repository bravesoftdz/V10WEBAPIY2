{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPIECERGCAU ()
Mots clefs ... : TOF;BTPIECERGCAU
*****************************************************************}
Unit BTPIECERGCAU_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,fe_main,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     SaisUtil,
     HTB97,
     graphics,
     Math,
     BTPUtil,
     AglInit;

Type
  TOF_BTPIECERGCAU = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TOBPIECERG : TOB;
    NumCaution : Thedit;
    BanQue : ThValComboBox;
    MtCaution : ThNumedit;
    NumeroRib : THEdit;
    DEV : RDevise;
    IsCompta : boolean;
    Auxiinterv : string;
    //
    procedure ChangeBanque(Sender: Tobject);
    procedure ChangeNumCaution(Sender: Tobject);
    procedure ChangeMtCaution (Sender:TObject);
    procedure ChangeRib (Sender : TObject);
    procedure RIBElipsisClick (Sender : TObject);
		procedure reinitcaution;
    procedure recalculeCaution;
  end ;

Implementation
uses Paramsoc,FactCalc,TiersUtil;

procedure TOF_BTPIECERGCAU.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCAU.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCAU.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCAU.OnLoad ;
begin
  Inherited ;
  if Auxiinterv <> '' then
  begin
  	ChangeRib(ecran);
		ChangeNumCaution (ecran);
  end;
end ;

procedure TOF_BTPIECERGCAU.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBPIECERG := LaTOB;
  IsCompta := (not GetParamSocSecur('SO_GCDESACTIVECOMPTA',True));
  Auxiinterv :=  TOBPieceRG.getString('PRG_FOURN');
  DEV.code := LaTOB.GetString('DEV');
  if Auxiinterv='' then
  begin
		THEdit(GetControl('PRG_NUMERORIB')).visible := false;
    THlabel(GetControl('DESIGNATIONRIB')).Visible := false;
    THlabel(GetControl('NOMBANQUE')).Visible := false;
    THValComboBox (GetControl('PRG_BANQUECP')).visible := true;
  end else
  begin
		NumeroRib := THEdit(GetControl('PRG_NUMERORIB'));
    NumeroRIB.visible := true;
    NumeroRib.OnExit := ChangeRib;
    NumeroRib.OnElipsisClick := RIBElipsisClick;
    THlabel(GetControl('NOMBANQUE')).Visible := true;
    THlabel(GetControl('NOMBANQUE')).caption := '';
//		THEdit(GetControl('PRG_NUMERORIB')).Plus := ' AND R_AUXILIAIRE="'+TiersAuxiliaire (Auxiinterv,False)+'"';
    THlabel(GetControl('DESIGNATIONRIB')).Visible := true;
    THlabel(GetControl('DESIGNATIONRIB')).caption := '';
    THValComboBox (GetControl('PRG_BANQUECP')).visible := false;
  end;
  BanQue := THValComboBox (GetControl('PRG_BANQUECP'));
  if IsCompta then
  begin
    // Code Banque pour la caution banquaire
    BanQue.OnChange := ChangeBanque;
    // Creation d'une nouvelle banque
  end;
  // Modif du Numero de caution
  Numcaution:= THedit(Getcontrol('PRG_NUMCAUTION'));
  Numcaution.onexit := ChangeNumcaution;
  // Modif Montant de caution
  MtCaution := THNumEdit(GetControl('PRG_CAUTIONMTDEV'));
  MTcaution.OnExit := ChangeMtCaution;

  GetInfosDevise(DEV) ;
  SetControlText ('DEVISE',DEV.Symbole );
end ;

procedure TOF_BTPIECERGCAU.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCAU.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCAU.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCAU.ChangeBanque(Sender: Tobject);
begin
  if Banque.Value = '' then
  begin
    SetControlText ('PRG_NUMCAUTION','');
    SetControlEnabled ('PRG_NUMCAUTION',false);
    SetControlProperty ('PRG_NUMCAUTION','color',clInactiveCaption);
    reinitcaution;
  end else
  begin
    SetControlEnabled ('PRG_NUMCAUTION',true);
    SetControlProperty ('PRG_NUMCAUTION','color',clWindow);
    SetControlProperty('PRG_CAUTIONMTDEV','value',ceil(TOBPieceRG.GetDouble('PRG_MTTTCRGDEV')));
    recalculeCaution;
  end;

end;

procedure TOF_BTPIECERGCAU.ChangeMtCaution(Sender: TObject);
begin
	recalculeCaution;
end;

procedure TOF_BTPIECERGCAU.ChangeNumCaution(Sender: Tobject);
begin
  if (Numcaution.Text  = '') then
  begin
  	Setcontrolenabled ('PRG_CAUTIONMTDEV',true);
  end else
  begin
  	Setcontrolenabled ('PRG_CAUTIONMTDEV',false);
  end;

  if not isCompta then
  begin
    SetControlProperty('PRG_CAUTIONMTDEV','value',ceil(TOBPieceRG.GetDouble('PRG_MTTTCRGDEV')));
    recalculeCaution;
  end;
end;

procedure TOF_BTPIECERGCAU.reinitcaution;
begin
  SetControlProperty('PRG_CAUTIONMTDEV','value',0);
  SetControlProperty('PRG_CAUTIONMT','value',0);
end;

procedure TOF_BTPIECERGCAU.recalculeCaution;
var XP : double;
begin
  CalculeMontantsAssocie (MtCaution.value,XP,DEV);
  SetControlProperty('PRG_CAUTIONMT','value',XP);
end;

procedure TOF_BTPIECERGCAU.ChangeRib(Sender: TObject);
var RIB : string;
		CodeBQ,CodeTiers,CodeAuxi : string;
    OkBQ : boolean;
begin
  CodeTiers :=  TOBPieceRG.getString('PRG_FOURN');
  CodeAuxi := TiersAuxiliaire (CodeTIers,false,'FOU');
  OkBQ := GetCodeBQ (CodeAuxi, StrToInt(THEdit(GetControl('PRG_NUMERORIB')).Text),CodeBQ);
  if not OKBQ then
  begin
    THlabel(GetControl('DESIGNATIONRIB')).caption := '';
    SetControlText ('PRG_NUMCAUTION','');
    SetControlEnabled ('PRG_NUMCAUTION',false);
    SetControlProperty ('PRG_NUMCAUTION','color',clInactiveCaption);
    THlabel(GetControl('NOMBANQUE')).caption := '';
    reinitcaution;
  end else
  begin
		RIB :=  GetRIBParticulier (CodeAuxi,StrToInt(THEdit(GetControl('PRG_NUMERORIB')).Text));
    THlabel(GetControl('DESIGNATIONRIB')).caption := RIB;
    SetControlEnabled ('PRG_NUMCAUTION',true);
    SetControlProperty ('PRG_NUMCAUTION','color',clWindow);
    SetControlProperty('PRG_CAUTIONMTDEV','value',ceil(TOBPieceRG.GetDouble('PRG_MTTTCRGDEV')));
    THlabel(GetControl('NOMBANQUE')).Caption := CodeBQ;
    recalculeCaution;
  end;
end;

procedure TOF_BTPIECERGCAU.RIBElipsisClick(Sender: TObject);
var CodeTiers,sNumeroRib,NatureAuxi,CodeAuxi : string;
begin
  CodeTiers :=  TOBPieceRG.getString('PRG_FOURN');
  CodeAuxi := TiersAuxiliaire (CodeTIers,false,'FOU');

  NatureAuxi:='FOU';
  //appel de la fiche RIB pour récupération code banque...
  sNumeroRIB := AGLLanceFiche('YY','YYRIB',CodeAuxi,'','NumAux='+ TiersAuxiliaire (CodeTIers,false,NatureAuxi));
  if sNumeroRib <> '' then
  begin
    THEdit(GetControl('PRG_NUMERORIB')).Text := sNumeroRib;
    ChangeRib(Self);
  end;
end;

Initialization
  registerclasses ( [ TOF_BTPIECERGCAU ] ) ;
end.
