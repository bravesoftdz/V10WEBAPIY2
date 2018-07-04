{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/11/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : BTPIECERG ()
Mots clefs ... : TOF;BTPIECERG
*****************************************************************}
Unit UTOFBTPIECERG ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF,UTOB,SaisUtil,M3FP,ENt1,FacTRg,factcalc,
{$IFDEF EAGLCLIENT}
      MaineAgl,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,Fe_Main, MajTable,
      banquecp,
{$ENDIF}
      Graphics,Dicobtp,HTB97,ENtGC,AGlInit,Math,ParamSoc;
Type
  TOF_BTPIECERG = Class (TOF)
  private
    TOBPiece,TOBPieceRG,TOBBases,TOBPorcs : TOB;
    TauxRg : ThNumedit;
    TypeRg : THValComboBox;
    NumCaution : Thedit;
    BanQue : ThValComboBox;
    NewBanque : TToolbarbutton97;
    MtCaution : ThNumedit;
    DEV : RDevise;
    PorcHt, PorcTTC : double;
    IsCompta : boolean;
    ValHT,ValTTC : double;
    procedure ChangeTaux(Sender: Tobject);
    procedure ChangeType(Sender: Tobject);
    procedure ChangeBanque(Sender: Tobject);
    procedure ClickParamBanque(Sender: Tobject);
    procedure ChangeNumCaution(Sender: Tobject);
    procedure recalculeCaution;
    procedure reinitcaution;
    procedure ChangeMtCaution (Sender:TObject);
    procedure BCOTRAITClick (Sender : Tobject);
    procedure MtTTCRGChange (Sender : TObject);
    function GetMax: double;
    procedure MtHTRGChange(Sender: TObject);
    procedure MTMANUELCHECKED (Sender : Tobject);
  public
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

  const
	TexteMsgCompl: array[1..2] of string 	= (
          {1}        'Vous devez renseigner un type d''application',
          {2}        'Vous devez renseigner un Taux d''application'
            );

Implementation
uses UCotraitance,UspecifPOC;

procedure TOF_BTPIECERG.ChangeMtCaution (Sender:TObject);
begin
  recalculeCaution;
end;

procedure TOF_BTPIECERG.reinitcaution;
begin
  SetControlProperty('PRG_CAUTIONMTDEV','value',0);
  SetControlProperty('PRG_CAUTIONMT','value',0);
end;

procedure TOF_BTPIECERG.recalculeCaution;
var XP,XC,XE : double;
    SaisieContre : boolean;
begin
  SaisieContre := (laTob.getValue('PRG_SAISIECONTRE')='X');
  CalculeMontantsAssocie (MtCaution.value,XP,DEV);
  SetControlProperty('PRG_CAUTIONMT','value',XP);
end;

procedure TOF_BTPIECERG.OnDelete ;
begin
  Inherited ;
  latob.initvaleurs;
  if (LaTOB <> Nil) then TheTob := LaTOB ;
  ecran.close;
end ;

procedure TOF_BTPIECERG.OnUpdate ;
begin
  Inherited ;

  if TypeRg.Value = '' then
   begin
   PGIBoxAF(TexteMsgCompl[1],Ecran.Caption);
   SetFocusControl('PRG_TYPERG') ; TForm(Ecran).ModalResult:=0;
   Exit;
   end;
  if TauxRg.Value = 0 then
   begin
   PGIBoxAF(TexteMsgCompl[2],Ecran.Caption);
   SetFocusControl('PRG_TAUXRG') ; TForm(Ecran).ModalResult:=0;
   Exit;
   end;
  if ((IsCompta) and ((Banque.value = '') or (NumCaution.text = ''))) or
   (( not IsCompta) and (NumCaution.text = '')) then
   begin
   laTob.putvalue('PRG_CAUTIONMTDEV',0);
   laTob.putvalue('PRG_CAUTIONMT',0);
   end;
  if (LaTOB <> Nil) then TheTob := LaTOB ;
end ;

procedure TOF_BTPIECERG.OnLoad ;
var ReliqCaution : Double;
begin
  Inherited ;

  PorcHt := CalculPort (true,TOBPorcs);
  PorcTTC := CalculPort (false,TOBPorcs);
  //
  ValHT := TOBPiece.getvalue('GP_TOTALHTDEV') - PorcHt;
  ValTTC := TOBPiece.getvalue('GP_TOTALTTCDEV') - PorcTTC;
  //
  if LaTob.getValue('PRG_TYPERG')='' then
   begin
    SetControlProperty ('PRG_TYPERG','value','TTC');
   end;
  SetControlProperty('GP_TOTALHTDEV','value',ValHT);
  SetControlProperty('GP_TOTALTTCDEV','value',ValTTC);
  if latob.getValue('PRG_TAUXRG') <> 0 then Setcontrolvisible('Bdelete',true);
  if LaTob.getValue('PRG_TAUXRG')= 0 then
   begin
    if VH_GC.BTCODESPECIF = '001' then
    begin
      SetControlProperty ('PRG_TAUXRG','value',GetTauxAffairePOC(TOBPiece.GetString('GP_AFFAIRE'),TOBPiece.GetDateTime('GP_DATEPIECE')));
    end else
    begin
   SetControlProperty ('PRG_TAUXRG','value',5);
    end;
   ChangeTaux (self);
   end;
end ;

function TOF_BTPIECERG.GetMax : double ;
var sens : double;
		Value : double;
begin
  if ValTTC < 0 then sens := -1 else Sens := 1;
  Value := Abs(ValTTC);
	result := Arrondi(ceil(value*(TauxRg.Value/100)),0) * sens;
end;

procedure TOF_BTPIECERG.OnArgument (S : String ) ;
begin
  Inherited ;

  IsCompta := (not GetParamSocSecur('SO_GCDESACTIVECOMPTA',True));
  // Type de retenu de garantie
  TypeRg := THValComboBox (GetControl('PRG_TYPERG'));
  TypeRg.OnChange := ChangeType;
  // Pourcentage de retenue
  TauxRg := THNumEdit(GetControl('PRG_TAUXRG'));
  TauxRg.OnExit  := ChangeTaux;

  BanQue := THValComboBox (GetControl('PRG_BANQUECP'));
  NewBanque := TToolbarButton97 (getcontrol('BANQUEPARAM'));
  if IsCompta then
  begin
    // Code Banque pour la caution banquaire
    BanQue.OnChange := ChangeBanque;
    // Creation d'une nouvelle banque
    NewBanque.OnClick := ClickParamBanque;
  end;
  // Modif du Numero de caution
  Numcaution:= THedit(Getcontrol('PRG_NUMCAUTION'));
  Numcaution.onexit := ChangeNumcaution;
  // Modif Montant de caution
  MtCaution := THNumEdit(GetControl('PRG_CAUTIONMTDEV'));
  MTcaution.OnExit := ChangeMtCaution;
  // Mise en place de TOBPieceRG
  TOBPieceRG := TOB(LaTob);
  // Mise en place de TOBPiece
  TOBPiece := TOB(LaTob.Data );
  if Pos(TOBPiece.getString('GP_NATUREPIECEG'),'FBT;FF')>0 then
  begin
    THEdit(GetControl('PRG_MTMANUEL')).visible := true;
    THedit(Getcontrol('PRG_MTTTCRGDEV')).OnExit := MtTTCRGChange;
    THedit(Getcontrol('PRG_MTHTRGDEV')).OnExit := MtHTRGChange;
    THEdit(GetControl('PRG_MTMANUEL')).OnClick := MTMANUELCHECKED;
  end else
  begin
     THEdit(GetControl('PRG_MTMANUEL')).visible := false;
  end;
  DEV.Code:=TOBPiece.GetValue('GP_DEVISE');
  // Mise en place de TOBBases
  TOBBases := TOB(TOBPiece.data);
  // Mise en Place des TOPorcs
  TOBPorcs := TOB(TOBBases.data);
  if (DEV.Code = '') then DEV.Code:=V_PGI.DevisePivot ;
  GetInfosDevise(DEV) ;
  SetControlText ('PRG_DEVISE',DEV.Symbole );
  SetControlText ('PRG_DEVISE1',DEV.Symbole );
  SetControlText ('PRG_DEVISE2',DEV.Symbole );
  if (not IsCompta) then
  begin
    Banque.visible := false;
    NewBanque.visible := false;
    SetControlEnabled ('PRG_NUMCAUTION',true);
//    SetControlProperty ('PRG_NUMCAUTION','color',clWindow);
    SetControlProperty('PRG_CAUTIONMTDEV','value',getMax);
    if not TOBPieceRg.getBoolean('PRG_MTMANUEL') then
    begin
    recalculeCaution;
  end;
  end;

  if TToolbarButton97 (getcontrol('BCOTRAIT')) <> nil then
  begin
    if not PieceTraitUsable (TOBPiece,nil) then SetControlVisible ('BCOTRAIT',false)
                                           else TToolbarButton97 (getcontrol('BCOTRAIT')).onClick := BCOTRAITClick;
  end;
end ;

procedure TOF_BTPIECERG.ChangeNumCaution (Sender:Tobject ) ;
begin
  (*
  if (Numcaution.Text  = '') then
  begin
  	Setcontrolenabled ('PRG_CAUTIONMTDEV',true);
  end else
  begin
  	Setcontrolenabled ('PRG_CAUTIONMTDEV',false);
  end;
  *)
  if not isCompta then
  begin
    SetControlProperty('PRG_CAUTIONMTDEV','value',Getmax);
    recalculeCaution;
  end;
  if (Numcaution.Text <> '' ) and (BanQue.Value<>'') then
  begin
    Setcontrolenabled ('PRG_BANQUECP',false);
    Setcontrolenabled ('PRG_CAUTIONMTDEV',false);
  end else
  begin
    Setcontrolenabled ('PRG_BANQUECP',true);
    Setcontrolenabled ('PRG_CAUTIONMTDEV',true);
  end;
end;

procedure TOF_BTPIECERG.ClickParamBanque (Sender:Tobject ) ;
begin
//FicheBanqueCP ('',taModif,0);
end;

procedure TOF_BTPIECERG.ChangeTaux (Sender:Tobject ) ;
var XP,XE : double;
    SaisieContre : boolean;
begin
  SaisieContre := (laTob.getValue('PRG_SAISIECONTRE')='X');
  // Partie HT
  SetControlProperty ('PRG_MTHTRGDEV','value',arrondi(ValHT*(TauxRg.Value /100),DEV.decimale));
  CalculeMontantsAssocie (THNumEdit(GetControl('PRG_MTHTRGDEV')).value,XP,DEV);
  SetControlProperty('PRG_MTHTRG','value',XP);
  // Partie TTC
  SetControlProperty('PRG_MTTTCRGDEV','value',arrondi(ValTTC*(TauxRg.Value/100),DEV.decimale));
  CalculeMontantsAssocie (THNumEdit(GetControl('PRG_MTTTCRGDEV')).value,XP,DEV);
  SetControlProperty('PRG_MTTTCRG','value',XP);
  //
  THCheckBox(getControl('PRG_MTMANUEL')).Checked := false;
  // caution bancaire
  if ((IsCompta) and  (Banque.Value <> '') and (numcaution.Text='')) or
   ((not IsCompta) and (NumCaution.text='')) then
   begin
   SetControlProperty('PRG_CAUTIONMTDEV','value',GetMax);
   recalculeCaution;
   end;
end;

procedure TOF_BTPIECERG.ChangeType (Sender:Tobject ) ;
begin
  if TypeRG.Value ='TTC' then
   begin
   SetControlVisible ('GP_TOTALHTDEV',false);
   SetControlVisible ('PRG_MTHTRGDEV',false);
   SetControlVisible ('PRG_MTHTRG',false);

   SetControlVisible ('GP_TOTALTTCDEV',true);
   SetControlVisible ('PRG_MTTTCRGDEV',true);
   SetControlVisible ('PRG_MTTTCRG',true);
   end else
   begin
   SetControlVisible ('GP_TOTALHTDEV',true);
   SetControlVisible ('PRG_MTHTRGDEV',true);
   SetControlVisible ('PRG_MTHTRG',true);

   SetControlVisible ('GP_TOTALTTCDEV',false);
   SetControlVisible ('PRG_MTTTCRGDEV',false);
   SetControlVisible ('PRG_MTTTCRG',false);
   end;
end;

procedure TOF_BTPIECERG.ChangeBanque (Sender:Tobject ) ;
begin
  if Banque.Value = '' then
   begin
   SetControlText ('PRG_NUMCAUTION','');
   SetControlEnabled ('PRG_NUMCAUTION',false);
    //

    //   SetControlProperty ('PRG_NUMCAUTION','color',clbtnFace);
//    reinitcaution;
   end else
   begin
   SetControlEnabled ('PRG_NUMCAUTION',true);
    //   SetControlProperty ('PRG_NUMCAUTION','color',clWindow);
   SetControlProperty('PRG_CAUTIONMTDEV','value',GetMAx);
   recalculeCaution;
   end;
end;

procedure TOF_BTPIECERG.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERG.BCOTRAITClick(Sender: Tobject);
begin
	//
end;


procedure TOF_BTPIECERG.MtTTCRGChange(Sender: TObject);
var Ratio : Double;
    Xp,XD : double;
begin
  if not THCheckbox (GEtControl('PRG_MTMANUEL')).Checked then exit;
  ratio := ValHT/valTTC;
  CalculeMontantsAssocie (THNumEdit(GetControl('PRG_MTTTCRGDEV')).value,XP,DEV);
  SetControlProperty('PRG_MTTTCRG','value',XP);
  //
  SetControlProperty('PRG_MTHTRGDEV','value',arrondi(THNumEdit(GetControl('PRG_MTTTCRGDEV')).value*ratio,DEV.decimale));
  CalculeMontantsAssocie (THNumEdit(GetControl('PRG_MTHTRGDEV')).value,XP,DEV);
  SetControlProperty('PRG_MTHTRG','value',XP);
  //
  THNumEdit(GetControl('PRG_TAUXRG')).Value := arrondi(VALTTC / THNumEdit(GetControl('PRG_MTTTCRGDEV')).value,2);
end;

procedure TOF_BTPIECERG.MtHTRGChange(Sender: TObject);
var Ratio : Double;
    Xp,Xd : double;
begin
  if not THCheckbox (GEtControl('PRG_MTMANUEL')).Checked then exit;
  ratio := ValTTC/valHT;
  CalculeMontantsAssocie (THNumEdit(GetControl('PRG_MTHTRGDEV')).value,XP,DEV);
  SetControlProperty('PRG_MTHTRG','value',XP);
  // Partie TTC
  SetControlProperty('PRG_MTTTCRGDEV','value',arrondi(THNumEdit(GetControl('PRG_MTHTRGDEV')).value*ratio,DEV.decimale));
  CalculeMontantsAssocie (THNumEdit(GetControl('PRG_MTTTCRGDEV')).value,XP,DEV);
  SetControlProperty('PRG_MTTTCRG','value',XP);
  //
  THNumEdit(GetControl('PRG_TAUXRG')).Value := arrondi(VALHT / THNumEdit(GetControl('PRG_MTHTRGDEV')).value,2);
end;

procedure TOF_BTPIECERG.MTMANUELCHECKED(Sender: Tobject);
begin
  if THCheckbox (GEtControl('PRG_MTMANUEL')).Checked then
  begin
    THNumEdit (getControl('PRG_MTTTCRGDEV')).enabled := true;
    THNumEdit(getControl('PRG_MTHTRGDEV')).enabled := true;
  end else
  begin
    THNumEdit(getControl('PRG_MTTTCRGDEV')).enabled := false;
    THNumEdit(getControl('PRG_MTHTRGDEV')).enabled := false;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTPIECERG ] ) ;
end.

