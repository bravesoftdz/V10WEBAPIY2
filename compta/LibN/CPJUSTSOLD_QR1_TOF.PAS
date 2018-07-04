{***********UNITE*************************************************
Auteur  ...... : Frédéric BELLARD
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPJUSTSOLD ()
Mots clefs ... : TOF;CPJUSTSOLD
*****************************************************************}
Unit CPJUSTSOLD_QR1_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     eQRS1,MaineAGL, UTob,
{$else}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,     // AGLLanceFiche
     QRS1,
 {$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Saisutil,
     TofMeth, {JP 30/06/06}
     ParamSoc,		//GetParamSocSecur YMO
     Ent1;

procedure CPLanceFiche_JustSold(Origine : TSuiviMP = SmpAucun);

Type
  TOF_CPJUSTSOLD_QR1 = Class (TOF_Meth)
    private
      TypEdit : THValComboBox;
      TypEcr : THValComboBox;
      NatComp : THValComboBox;
      EnSitu  : TCheckBox;
      TypLet  : THValComboBox;
      ExCompSold : TCheckBox;
      ContrLett : TCheckBox;
      ChoixTabLibres : THRadioGroup;
      LAuxDeb : THLabel;
      CptAuxDeb : THEdit;
      LAuxFin : THLabel;
      CptAuxFin : THEdit;
      LGenDeb : THLabel;
      CptGenDeb : THEdit;
      LGenFin : THLabel;
      CptGenFin : THEdit;
      OnTiers : Boolean ;
      AffichageMonnaie : THRadioGroup;
      DatComp : THEdit;
      Devise : THValComboBox;
      SorteSens : THValComboBox;
      CptAssoc : TCheckBox;
      AffichCol : TCheckBox;
      IndTypEdit : THEdit;
      TiersPayeur : THEdit;
      TriPar : THEdit;
      TriLib : TCheckBox;
      SepMon : TCheckBox;
      SepCom : TCheckBox;
      SepEGen : TCheckBox;
      SepPGen : TCheckBox;
      SepRup : TCheckBox;
      SepTie : TCheckBox;
      SepPag : TCheckBox;
      Rupture1 : THEdit;
      Rupture2 : THEdit;
      ConditionSQLEcrS : THEdit;
      ConditionSQLEcrF : THEdit;
      ConditionSQLEcrW : THEdit;
      ConditionSQLEcrO : THEdit;
      ConditionSQLCptS : THEdit;
      ConditionSQLCptF : THEdit;
      ConditionSQLCptW : THEdit;
      ConditionSQLCptO : THEdit;  
      DeviseAff        : RDevise; // FQ 20047

      procedure TypEditOnChange(Sender: TObject);
      procedure EnSituOnClick(Sender: TObject);
      procedure DatCompOnChange (Sender : TObject) ;
      procedure NatCompOnChange (Sender : TObject) ;
      function GetTabletteColl(stNatureTiers : String): String;
      function GetTabletteTiers(stNatureTiers: String): String;
      procedure TypEcrOnChange(Sender: TObject);

      procedure EtablissementOnChange (Sender : TObject);
      procedure DeviseOnChange (Sender : TObject) ;
      procedure AffichageMonnaieOnClick (Sender : TObject);
      procedure SorteSensOnChange (Sender : TObject);

      procedure Rupture1OnDblClick (Sender : TObject);
      procedure TriParOnChange (Sender : TObject);
      procedure TriParOnDblClick (Sender : TObject);
      procedure ChoixTabLibresOnClick(Sender: TObject);
      procedure CptAssocOnClick(Sender: TObject);
      procedure UpdateOptions(vBoAvecRupt: Boolean);
      function  GetMinMaxCompte(stTable, stCol, stDefaut : String) : String;
      function GetWhereColl(stNatureTiers : String): String;
      function  GetWhereGenTie(stNatureGene : String): String; //fb 05/06/2006 FQ18023
      procedure GenererConditionSQL;
      procedure SqlEcr;
      procedure SqlCpt;
      procedure CompteOnExit(Sender : TObject) ; // FQ 20167
    public
      procedure OnArgument (S : String ) ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure AuxiElipsisClick         (Sender: TObject);

  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  TofTLTiersTri, TofTLTiersRupt, UTofMulParamGen,uLibWindows;

var
  OrigineLocal : TSuiviMP;
  cTypeGeneraux : string = 'GEN';
  cTypeTiers    : string = 'TIE';

procedure CPLanceFiche_JustSold(Origine : TSuiviMP = smpAucun);
begin
  OrigineLocal:=Origine;
  // Etat non utilisable sous ORACLE7
  if V_PGI.Driver = dbORACLE7 then
    PGIInfo( 'Cet état n''est pas utilisable sous Oracle 7. Veuillez utiliser l''état correspondant dans le menu Autres éditions' , 'Justificatif de soldes' )
  else
    AGLLanceFiche('CP', 'CPJUSTSOLD_QR1', '', '', '') ;
end;

procedure TOF_CPJUSTSOLD_QR1.TypEditOnChange(Sender: TObject);
var
  NomFiltre : String ;
begin
  inherited;

  OnTiers:=(TypEdit.Value = cTypeTiers);

  LGenDeb.Visible:=OnTiers;
  CptGenDeb.Visible:=OnTiers;
  CptGenFin.Visible:=OnTiers;
  LGenFin.Visible:=OnTiers;
  SepEGen.Visible:=OnTiers;
  SepPGen.Visible:=OnTiers;
  AffichCol.Visible:=OnTiers;
  { FQ 19755 BVE 19.04.07
  if TypEdit.ItemIndex=0 then begin
    END FQ 19755 }
  if TypEdit.Value = cTypeGeneraux then
  begin
    LAuxDeb.Caption:=TraduireMemoire('Comptes généraux de');
    NatComp.DataType:='ttNatGeneTIDTIC';
    NatComp.ItemIndex:=0;
  end
  else
  begin
    LAuxDeb.Caption:=TraduireMemoire('Comptes auxiliaires de');
    NatComp.DataType:='ttNatTiersCpta';
    NatComp.ItemIndex:=0;
  end;

{$IFDEF CCMP}
    if OrigineLocal=smpEncTous then
      NomFiltre:='QRJUSOLDETIEREC'
    else
      NomFiltre:='QRJUSOLDETIERDE' ;
{$ELSE}
    NomFiltre:='QRJUSOLDETIER' ;
{$ENDIF}
{$IFDEF CCMP}
    NatComp.Vide := false;

    if (OnTiers) then begin
      NatComp.enabled := true;
      if (OrigineLocal = smpEncTous) then begin
        NatComp.Datatype := 'TTNATTIERSCPTAENC';   // encaissement
        NatComp.ItemIndex:=0
       end
      else begin
        NatComp.Datatype := 'TTNATTIERSCPTADEC';  // decaissement
        NatCompOnChange(NatComp);
        NatComp.ItemIndex:=0
      end;
     end
    else begin
      if (OrigineLocal = smpEncTous) then
        NatComp.Value := 'TID'  // encaissement
      else
        NatComp.Value := 'TIC'; // decaissement
      NatCompOnChange(NatComp);
      NatComp.enabled := false;
      end;
{$ENDIF}
  NatCompOnChange(NatComp);
  SetControlText('INDTYPEDIT',IntToStr(TypEdit.itemindex));
end;

procedure TOF_CPJUSTSOLD_QR1.EnSituOnClick(Sender: TObject);
begin
  inherited;
  { FQ 19799 BVE 19.04.07
  if EnSitu.Checked then begin
    ExCompSold.Checked:=false;
    ExCompSold.Enabled:=False;
    TypLet.ItemIndex:=0;
   end
  else
    ExCompSold.Enabled:=true;
    END FQ 19799 }

  TypLet.Enabled:=Not EnSitu.Checked;
  ContrLett.Enabled:=EnSitu.Checked;
end;

procedure TOF_CPJUSTSOLD_QR1.DatCompOnChange (Sender : TObject) ;
begin
  if not IsValidDate(GetControlText('DATCOMP')) then
    DatComp.Text:=DateToStr(V_PGI.DateEntree) ;
end;

procedure TOF_CPJUSTSOLD_QR1.NatCompOnChange (Sender : TObject) ;
var
  CGen, CAux : string ;
begin
  inherited ;

  SetControlText('CPTAUX','');
  SetControlText('CPTAUX_','');
  SetControlText('CPTGEN','');
  SetControlText('CPTGEN_','');
{b fb 02/06/2006 FQ 18023}
  if (OnTiers) then begin
    CAux := GetTabletteTiers(NatComp.value);
    CGen := GetTabletteColl(NatComp.value)
   end
  else
    CAux := 'TZGTIDTIC';
{e fb 02/06/2006 FQ 18023}
  SetControlProperty('CPTAUX','DATATYPE',CAux);
  SetControlProperty('CPTAUX_','DATATYPE',CAux);
  SetControlProperty('CPTGEN','DATATYPE',CGen);
  SetControlProperty('CPTGEN_','DATATYPE',CGen);
end;

function TOF_CPJUSTSOLD_QR1.GetTabletteColl(stNatureTiers : String): String;
begin
  if stNatureTiers='CLI' then result := 'TZGCOLLCLIENT'
  else if stNatureTiers='FOU' then result := 'TZGCOLLFOURN'
  else if stNatureTiers='SAL' then result := 'TZGCOLLSALARIE'
  else if stNatureTiers='DIV' then result := 'TZGCOLLDIVERS'
  else if stNatureTiers='AUC' then result := 'TZGCOLLTOUTDEBIT'
  else if stNatureTiers='AUD' then result := 'TZGCOLLTOUTCREDIT'
  else result := 'TZGCOLLECTIF';
end;

function TOF_CPJUSTSOLD_QR1.GetTabletteTiers(stNatureTiers: String): String;
begin
  if stNatureTiers='CLI' then result := 'TZTCLIENT'
  else if stNatureTiers='FOU' then result := 'TZTFOURN'
  else if stNatureTiers='SAL' then result := 'TZTSALARIE'
  else if stNatureTiers='DIV' then result := 'TZTDIVERS'
  else if stNatureTiers='AUC' then result := 'TZTTOUTDEBIT'
  else if stNatureTiers='AUD' then result := 'TZTTOUTCREDIT'
  else if stNatureTiers='TID' then result := 'TZGTIDTIC'
  else if stNatureTiers='TIC' then result := 'TZGTIDTIC'
  else result := 'TZTTOUS';
end;

procedure TOF_CPJUSTSOLD_QR1.TypEcrOnChange(Sender: TObject);
begin
  inherited;
  if (TypEcr.Value='PRE')or(TypEcr.Value='SSI') then begin
    EnSitu.State:=cbUnchecked;
    EnSitu.Enabled:=False;
    TypLet.Value:='NL';
    TypLet.Enabled:=False;
   end
  else begin
    if EnSitu.Checked then begin
      TypLet.Enabled:=false;
      TypLet.Value:=TraduireMemoire('<<Tous>>');
     end
    else
      TypLet.Enabled:=true;
    EnSitu.Enabled:=True;
    end;
end;

procedure TOF_CPJUSTSOLD_QR1.EtablissementOnChange (Sender : TObject);
begin
  if ComboEtab.ItemIndex = 0 then
    SetControlText('ETABL','%')
  else
    SetControlText('ETABL',GetControlText('ETABLISSEMENT'));
end;

procedure TOF_CPJUSTSOLD_QR1.DeviseOnChange (Sender : TObject);
begin
  if Devise.ItemIndex=0 then begin
    AffichageMonnaie.ItemIndex := 0;
    AffichageMonnaie.Enabled := False;
   end
  else
    AffichageMonnaie.Enabled := True;
end;

procedure TOF_CPJUSTSOLD_QR1.AffichageMonnaieOnClick (Sender : TObject);
begin
  if AffichageMonnaie.ItemIndex=0 then
  { FQ 20047 BVE 25.04.07 }
    //SetControlText('DEVAFFICHAGE','Euro')
    SetControlText('DEVAFFICHAGE',DeviseAff.Libelle)
  { END FQ 20047 }
  else
    SetControlText('DEVAFFICHAGE','Devise');
end;

procedure TOF_CPJUSTSOLD_QR1.SorteSensOnChange (Sender : TObject);
begin
  if SorteSens.ItemIndex=0 then
    SetControlText('SENS','%')
  else
    SetControlText('SENS',GetControlText('SORTESENS'));
end;

procedure TOF_CPJUSTSOLD_QR1.Rupture1OnDblClick (Sender : TObject);
var
  St,St1,St2,St3,Arg : String ;
  i       : integer ;
  OkAssoc : integer ;
begin
  OkAssoc := 0;
  Arg     := GetControlText('RUPTURE1')+'|'+GetControlText('RUPTURE2');
  if GetControlVisible('CHOIXTABLIBRES') then
    Arg := Arg + '|' + GetControlText('CHOIXTABLIBRES')
  else
    Arg:= Arg + '|AUX';
  St      := CPLanceFiche_TLTIERSRUPT(Arg);
  St1     := ReadTokenPipe(St,'|');
  SetControlText('RUPTURE1',St1);
  SetControlText('RUPTURE2',St);
  for i:=0 to 9 do begin
    St2:=ReadTokenSt(St1);
    if (St2<>'#') and (St2<>'-') and (St2<>'') then begin
      if ChoixTabLibres.Value = 'AUX' then
        St3:=St3+'T0'+IntToStr(i)+';'
      else if ChoixTabLibres.Value = 'CLI' then
        St3:=St3+'LT'+IntToHex(i+1,1)+';'
      else
        St3 :=St3+'AR'+IntToHex(i+1,1)+';';
      OkAssoc:=OkAssoc+1 ;
      end;
    SetControlText('TRI',St3);
    end;

  // MAJ interface si aucune rupture sélectionnée
  UpdateOptions( OkAssoc > 0 ) ;
  if OkAssoc = 0 then
    begin
    SetControlText('RUPTURE1', '');
    SetControlText('RUPTURE2', '');
    SetControlText('TRI',   '');
    end ;
end;

procedure TOF_CPJUSTSOLD_QR1.TriParOnChange (Sender : TObject);
begin
  if GetControlText('TRI')='' then begin
    SetControlChecked('SAUTPAGE',False);
    SetControlChecked('SURRUPTURE',False);
   end
  else begin
    SetControlEnabled('SAUTPAGE',True);
    SetControlEnabled('SURRUPTURE',True);
    end;
end;

procedure TOF_CPJUSTSOLD_QR1.TriParOnDblClick (Sender : TObject);
var
  St,Arg:String;
begin
  Arg := GetControlText('TRI');
  St  := CPLanceFiche_TLTIERSTRI(Arg);
  SetControlText('TRI',St);
end;

procedure TOF_CPJUSTSOLD_QR1.ChoixTabLibresOnClick(Sender: TObject);
begin
  SetControlText('RUPTURE1','');
  SetControlText('RUPTURE2','');
  SetControlText('TRI','');
  SetControlEnabled('TRI',(ChoixTabLibres.Value <> 'SCL') and (ChoixTabLibres.Value <> 'CLI'));
end;

procedure TOF_CPJUSTSOLD_QR1.CptAssocOnClick(Sender: TObject);
begin
  inherited;
  if (OnTiers) then
    AffichCol.Visible:=true
  else begin
    AffichCol.Visible:=false;
    AffichCol.Checked:=false;
    end;
end;

procedure TOF_CPJUSTSOLD_QR1.OnArgument (S : String ) ;
BEGIN
  inherited ;
  { FQ 20047 BVE 25.04.07 }
  AffichageMonnaie := THRadioGroup(GetControl('AFFICHAGEMONNAIE'));
  DeviseAff.Code := GetParamSoc('SO_DEVISEPRINC');
  GetInfosDevise(DeviseAff);
  if DeviseAff.Libelle = '' then DeviseAff.Libelle := 'Euro';
  AffichageMonnaie.Items.Strings[0] := DeviseAff.Libelle;
  { END FQ 20047 }

  IndTypEdit := THEdit(GetControl('INDTYPEDIT'));
  TypEdit := THValComboBox(GetControl('TYPEDIT'));
  LAuxDeb := THLabel(GetControl('LAUXDEB'));
  CptAuxDeb := THEdit(GetControl('CPTAUX'));
  LAuxFin := THLabel(GetControl('LAUXFIN'));
  CptAuxFin := THEdit(GetControl('CPTAUX_'));
  LGenDeb := THLabel(GetControl('LGENDEB'));
  CptGenDeb := THEdit(GetControl('CPTGEN'));
  LGenFin := THLabel(GetControl('LGENFIN'));
  CptGenFin := THEdit(GetControl('CPTGEN_'));
  DatComp := THEdit(GetControl('DATCOMP'));
  EnSitu  := TCheckBox(GetControl('ENSITU'));
  NatComp := THValComboBox(GetControl('NATCOMP'));
  TypLet  := THValComboBox(GetControl('TYPLET'));    //
  ExCompSold := TCheckBox(GetControl('EXCOMPSOLD')); //
  TiersPayeur := THEdit(GetControl('TP'));
  TypEcr := THValComboBox(GetControl('TYPECR'));

  Devise := THValComboBox(GetControl('DEVISE'));
  SorteSens := THValComboBox(GetControl('SORTESENS'));


  Rupture1 := THEdit(GetControl('RUPTURE1'));
  Rupture2 := THEdit(GetControl('RUPTURE2'));
  TriPar := THEdit(GetControl('TRI'));
  ChoixTabLibres := THRadioGroup(GetControl('CHOIXTABLIBRES'));
  CptAssoc := TCheckBox(GetControl('CPTASSOCIES'));
  ConditionSQLEcrS := THEdit(GetControl('SQLECRS'));
  ConditionSQLEcrF := THEdit(GetControl('SQLECRF'));
  ConditionSQLEcrW := THEdit(GetControl('SQLECRW'));
  ConditionSQLEcrO := THEdit(GetControl('SQLECRO'));
  ConditionSQLCptS := THEdit(GetControl('SQLCPTS'));
  ConditionSQLCptF := THEdit(GetControl('SQLCPTF'));
  ConditionSQLCptW := THEdit(GetControl('SQLCPTW'));
  ConditionSQLCptO := THEdit(GetControl('SQLCPTO'));

  ContrLett := TCheckBox(GetControl('CONTRLETT')); //
  SepMon := TCheckBox(GetControl('SDMON'));   //
  SepCom := TCheckBox(GetControl('SDCOM'));   //
  SepEGen := TCheckBox(GetControl('SETGEN')); //
  SepPGen := TCheckBox(GetControl('SPGEN'));  //
  SepRup := TCheckBox(GetControl('SDRUP'));   //
  SepTie := TCheckBox(GetControl('SPTIE'));  //
  SepPag := TCheckBox(GetControl('LSDPAG'));     //
  TriLib := TCheckBox(GetControl('TRILIBELLE')); //
  AffichCol := TCheckBox(GetControl('AFFICHCOL')); //

  { FQ 20167 BVE 02.05.07 }
  CptAuxDeb.OnExit := CompteOnExit ;
  { END FQ 20167  }

  if TypEdit<>nil then begin
    TypEdit.OnChange := TypEditOnChange;
    { FQ 19755 BVE 19.04.07 }
    TypEdit.Value := cTypeTiers;
    { END FQ 19755 }
    if NatComp<>nil then
      NatComp.OnChange:=NatCompOnChange;
    TypEditOnChange(TypEdit);
    if IndTypEdit<>nil then
      IndTypEdit.Text := IntToStr(TypEdit.itemindex);
    end;

  if DatComp<>nil then begin
    DatComp.OnChange := DatCompOnChange;
    DatComp.Text:=DateToStr(V_PGI.DateEntree) ;
    end;

  if EnSitu<>nil then begin
    EnSitu.OnClick:=EnSituOnClick;
    EnSitu.Checked:=true;
    EnSitu.OnClick(nil);
    end;
  if NatComp<>nil then begin
    NatComp.OnChange:=NatCompOnChange;
    NatComp.Vide := true;
    NatComp.ItemIndex:=0;
    NatCompOnChange(nil);
    end;
  if TypEcr<>nil then begin
    TypEcr.OnChange:=TypEcrOnChange;
    TypEcr.ItemIndex := 0;
    TypEcrOnChange(nil);
    end;
  if TypLet<>nil then
    TypLet.ItemIndex:=0;
    
  if Tierspayeur<>nil then begin
    if (TypEdit.Value = cTypeTiers) and (VH^.OuiTP) then
      TiersPayeur.Text:='V'
    else
      TiersPayeur.Text:='F';
    SetControlText('TP', TiersPayeur.text);
    end;

  if ComboEtab <> nil then begin
    ComboEtab.OnChange:=EtablissementOnChange;
    ComboEtab.OnChange(ComboEtab);
    end;

  if Devise<>nil then begin
    Devise.OnChange:=DeviseOnChange;
    Devise.itemindex:=0;
    Devise.OnChange(NIL);
    end;

  if SorteSens<>nil then begin
    SorteSens.OnChange:=SorteSensOnChange;
    SorteSens.Value :='M';
    SorteSens.OnChange(nil);
    end;

  if AffichageMonnaie<>nil then begin
    { FQ 20047 BVE 25.04.07 }
    //SetControlText('DEVAFFICHAGE','Euro')
    SetControlText('DEVAFFICHAGE',DeviseAff.Libelle);
    { END FQ 20047 }
    AffichageMonnaie.OnClick:=AffichageMonnaieOnClick;
    AffichageMonnaie.ItemIndex:=0;
    end;

  if Rupture1<>nil then
    Rupture1.OnDblClick:=Rupture1OnDblClick;
  if Rupture2<>nil then
    Rupture2.OnDblClick:=Rupture1OnDblClick;
  if TriPar<>nil then begin
    TriPar.OnChange:=TriParOnChange;
    TriPar.OnDblClick:=TriParOnDblClick;
    SetControlEnabled('SAUTPAGE',False);
    SetControlEnabled('SURRUPTURE',False);
    end;
  if ChoixTabLibres<>nil then begin
    ChoixTabLibres.OnClick:=ChoixTabLibresOnClick;
    ChoixTabLibres.ItemIndex:=0;
    end;
  if CptAssoc<>nil then begin
    CptAssoc.OnClick := CptAssocOnClick;
    CptAssoc.Enabled:=false;
    CptAssoc.OnClick(nil);
    end;

  SetControlChecked('SDMON',true);
  SetControlChecked('SDCOM',true);
  SetControlChecked('SETGEN',true);
  SetControlChecked('SPGEN',true);
  SetControlChecked('SDRUP',true);
  SetControlChecked('SPTIE',true);
  SetControlChecked('LSDPAG',false);

  // Init options des ruptures
  UpdateOptions(false);
  ChoixTabLibres.Visible := TL_TIERSCOMPL_Actif;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    { FQ 20387 BVE 24.05.07
    THEdit(GetControl('T_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('T_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick; } 
    THEdit(GetControl('CPTAUX', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('CPTAUX_', true)).OnElipsisClick:=AuxiElipsisClick;
    { END FQ 20387 }
  end;

end;

procedure TOF_CPJUSTSOLD_QR1.UpdateOptions(vBoAvecRupt: Boolean);
// MAJ accès aux checkbox liés aux options de ruptures
begin
  if not vBoAvecRupt then begin
    SetControlChecked( 'SAUTPAGE',    False );
    SetControlChecked( 'SURRUPTURE',  False );
    end ;
  SetControlChecked( 'CPTASSOCIES', vBoAvecRupt );
  SetControlEnabled( 'CPTASSOCIES', vBoAvecRupt );
  SetControlEnabled( 'SAUTPAGE',    vBoAvecRupt );
  SetControlEnabled( 'SURRUPTURE',  vBoAvecRupt );
end;

procedure TOF_CPJUSTSOLD_QR1.OnUpdate ;
Var
  St,St1,St2,St3,St4,St5,S,S1,TriLibelle,Order,stTL : string ;
  i : integer;
  stIndice : string;
Begin
  S:='';
  S1:='';
  TriLibelle:='T_AUXILIAIRE' ;
  St:=GetControlText('TRI') ;
  St1:=GetControlText('RUPTURE1') ;
  St2:=GetControlText('RUPTURE2') ;
  if GetControlText('TRILIBELLE')='X' then
    TriLibelle:='T_LIBELLE'
  else
    TriLibelle:='T_AUXILIAIRE' ;

  // Construction condition sur ruptures tables libres
  // Codes utilisés pour la gestion des ruptures sur tables libres :
  //    # = non gérée
  //    - = gérée mais non sélectionnée
  //    * = tous les enregistrements de la table (sans fourchette de valeurs)
  //  sinon code de valeur début dans St3, fin dans St4
  if St1<>'' then
  begin
    { FQ 20392 BVE 25.05.07 }
    if (ChoixTabLibres.Value = 'CLI') or (ChoixTabLibres.Value = 'SCL') then
       SetControlText('SQLRUPTF','LEFT JOIN TIERSCOMPL ON YTC_AUXILIAIRE = E_AUXILIAIRE');
    { END FQ 20392 }
    for i:=0 to 9 do begin
      St3:=ReadTokenSt(St1) ;
      St4:=ReadTokenSt(St2) ;
      if (St3<>'-') and (St3>'#') and (St3<>'') then begin// uniquement les tables sélectionnées
        if (St3<>'*') then begin// Si on a une fourchette de valeurs
          if ChoixTabLibres.Value = 'AUX' then
            stTL := '( T_TABLE'+IntToStr(i)+'>="'+St3+'" AND T_TABLE'+IntToStr(i)+'<="'+St4+'" )'
          else if ChoixTabLibres.Value = 'CLI' then
            stTL := '( YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'>="'+St3+'" AND YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'<="'+St4+'" )'
          else if ChoixTabLibres.Value = 'SCL' then
            stTL := '( YTC_RESSOURCE'+IntToHex(i+1,1)+'>="'+St3+'" AND YTC_RESSOURCE'+IntToHex(i+1,1)+'<="'+St4+'" )';
          if not TCheckBox(GetControl('CPTASSOCIES')).Checked then begin
            if ChoixTabLibres.Value = 'AUX' then
              stTL := '( ' + stTL +' OR ( T_TABLE'+IntToStr(i)+'="" ) ) '
            else if ChoixTabLibres.Value = 'CLI' then
              stTL := '( ' + stTL +' OR ( YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'="" ) ) '
            else if ChoixTabLibres.Value = 'SCL' then
              stTL := '( ' + stTL +' OR ( YTC_RESSOURCE'+IntToHex(i+1,1)+'="" ) ) ' ;
            end;
          S := S + ' AND ' +  stTL ;
         end
        else // Sans fourchette de valeur
          if TCheckBox(GetControl('CPTASSOCIES')).Checked then begin
            if ChoixTabLibres.Value = 'AUX' then
              S := S + ' AND (T_TABLE'+IntToStr(i)+'<>"") '
            else if ChoixTabLibres.Value = 'CLI' then
              S := S + ' AND (YTC_TABLELIBRETIERS'+IntToHex(i+1,1)+'<>"") '
            else if ChoixTabLibres.Value = 'SCL' then
              S := S + ' AND (YTC_RESSOURCE'+IntToHex(i+1,1)+'<>"") ' ;
            end;
        end;
      end;
  end
  else            
    { FQ 20392 BVE 25.05.07 }
     SetControlText('SQLRUPTF','');    
    { END FQ 20392 }

  // conditions auxiliaire
{b fb 05/06/2006 FQ18023}
  if (OnTiers) then begin
    if (GetControlText('CPTAUX')='')  and (GetControlText('CPTAUX_')<>'') then
      S:=S+' AND T_AUXILIAIRE<="'+GetControlText('CPTAUX_')+'"' ;
    if (GetControlText('CPTAUX')<>'') and (GetControlText('CPTAUX_')<>'') then
      S:=S+' AND T_AUXILIAIRE>="'+GetControlText('CPTAUX')+'" AND T_AUXILIAIRE<="'+GetControlText('CPTAUX_')+'"'
   end
  else begin
    if (GetControlText('CPTAUX')='')  and (GetControlText('CPTAUX_')<>'') then
      S:=S+' AND G_GENERAL<="'+GetControlText('CPTAUX_')+'"' ;
    if (GetControlText('CPTAUX')<>'') and (GetControlText('CPTAUX_')<>'') then
      S:=S+' AND G_GENERAL>="'+GetControlText('CPTAUX')+'" AND G_GENERAL<="'+GetControlText('CPTAUX_')+'"'
    end;
{e fb 05/06/2006 FQ18023}

  // conditions collectif
  if (GetControlText('CPTGEN')='')  and (GetControlText('CPTGEN_')<>'') then
    S:=S+' AND T_COLLECTIF<="'+GetControlText('CPTGEN_')+'"' ;
  if (GetControlText('CPTGEN')<>'') and (GetControlText('CPTGEN_')<>'') then
    S:=S+' AND T_COLLECTIF>="'+GetControlText('CPTGEN')+'" AND T_COLLECTIF<="'+GetControlText('CPTGEN_')+'"' ;

  // conditions nature auxi
  if TComboBox(GetControl('NATCOMP')).itemindex<>0 then
    S:=S+' AND T_NATUREAUXI="'+GetControlText('NATCOMP')+'"' ;

  // Mise en place dans le champ caché
  SetControlText('WHERE',S) ;

  // MAJ pour gestion des ruptures dans l'état
  for i:=0 to 9 do begin
    SetControlText('T0'+IntToStr(i),'') ;
    SetControlText('ORDER'+IntToStr(i),'') ;
    end ;
  i := 0 ;
  while St<>'' do begin
    St5:=ReadTokenSt(St) ;
    if (St5<>'') then begin
        stIndice := copy(St5,3,1);
      if ChoixTabLibres.Value = 'AUX' then
        SetControlText('ORDER'+IntToStr(i),'T_TABLE' + stIndice)
      else if ChoixTabLibres.Value = 'CLI' then
        SetControlText('ORDER'+IntToStr(i),'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1))
      else if ChoixTabLibres.Value = 'SCL' then
        SetControlText('ORDER'+IntToStr(i),'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1));
      SetControlText('T0'+IntToStr(i),St5) ;
      if (St='') then begin
        if ChoixTabLibres.Value = 'AUX' then
          S1:=S1+'T_TABLE'+copy(St5,3,1)
        else if ChoixTabLibres.Value = 'CLI' then
          S1:=S1+'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1)
        else if ChoixTabLibres.Value = 'SCL' then
          S1:=S1+'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1)
       end
      else begin
        if ChoixTabLibres.Value = 'AUX' then
          S1:=S1+'T_TABLE'+copy(St5,3,1)+','
        else if ChoixTabLibres.Value = 'CLI' then
          S1:=S1+'YTC_TABLELIBRETIERS'+IntToHex(StrToInt(stIndice),1)+','
        else if ChoixTabLibres.Value = 'SCL' then
          S1:=S1+'YTC_RESSOURCE'+IntToHex(StrToInt(stIndice),1)+',';
        end;
      end ;
      inc(i) ;
    end ;

  if S1<>'' then
    Order:=S1+','+TriLibelle
  else
    Order:=TriLibelle ;

  { Mise à jour des requêtes pour récupérer les libellés des données des tables libres }
  for i := 0 to 9 do begin
    if ChoixTabLibres.Value = 'AUX' then
      SetControlText ('SQL'+IntToStr(i),'select nt_libelle FROM natcpte WHERE nt_typecpte="T0'+IntToStr(i)+'" and nt_nature=')
    else if ChoixTabLibres.Value = 'CLI' then
      SetControlText ('SQL'+IntToStr(i),'select yx_libelle FROM choixext WHERE yx_type="LT'+IntToHex(i+1,1)+'" and yx_code=')
    else if ChoixTabLibres.Value = 'SCL' then
      SetControlText ('SQL'+IntToStr(i),'select ars_libelle FROM ressource WHERE ars_ressource=');
    end;
  { FQ 20026 BVE 24.05.07 }
  TFQRS1(Ecran).WhereSQL:=TFQRS1(Ecran).WhereSQL + '|'+Order + ', E_GENERAL, E_ETATLETTRAGE DESC, '+
                          ' E_LETTRAGE, E_DATECOMPTABLE, E_DATEECHEANCE, E_NUMEROPIECE, ' +
                          ' E_NUMECHE ';
  { END FQ 20026 }
end ;

function TOF_CPJUSTSOLD_QR1.GetMinMaxCompte(stTable, stCol, stDefaut : String) : String;
var
  Q : TQuery;
  stWhere : String;
begin
  if stTable = 'GENERAUX' then begin
{b fb 05/06/2006 FQ 18023}
    if (OnTiers) then begin
      if NatComp.Value <> '' then
        stWhere := GetWhereColl(NatComp.Value)
      else
        stWhere := ' WHERE G_COLLECTIF="X"'
     end
    else begin
      if NatComp.Value <> '' then
        stWhere := GetWhereGenTie(NatComp.Value)
      else
        stWhere := ' WHERE G_NATUREGENE="TID" OR G_NATUREGENE="TIC" '
      end;
{e fb 05/06/2006 FQ 18023}
   end
  else begin
    if NatComp.Value <> '' then
      stWhere := ' WHERE T_NATUREAUXI="' + NatComp.Value + '"';
{b fb 05/06/2006 FQ 18023}
    {else
       stWhere := ' WHERE T_PAYEUR<>''';}
{e fb 05/06/2006 FQ 18023}
    end;
  Q := OpenSQL('SELECT ' + stCol +' CODE FROM ' + stTable + stWhere,True);
  if not Q.Eof then
    Result := Q.FindField('CODE').asString
  else
    Result := stDefaut;
  Ferme(Q);
end;

function TOF_CPJUSTSOLD_QR1.GetWhereColl(stNatureTiers : String): String;
begin
  result := ' WHERE G_COLLECTIF="X"';
  if stNatureTiers='CLI' then
    result := result + ' AND G_NATUREGENE="COC"'
  else if stNatureTiers='FOU' then
    result := result + ' AND G_NATUREGENE="COF"'
  else if stNatureTiers='SAL' then
    result := result + ' AND G_NATUREGENE="COS"'
  else if stNatureTiers='DIV' then
    result := result + ' AND G_NATUREGENE="COD"'
  else if stNatureTiers='AUC' then
    result := result + ' AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD")' 		else
  if stNatureTiers='AUD' then
    result := result + ' AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD")';
end;

procedure TOF_CPJUSTSOLD_QR1.OnLoad;
var
  RDev: RDevise;
begin
  inherited;
  SetControlText('SYMBOLEDEVISE',' ');
  if GetControlText('SYMBMON')='X' then begin
    if (AffichageMonnaie.ItemIndex = 0) or (AffichageMonnaie.Value=V_PGI.DevisePivot) then begin
      SetControlText('SYMBOLEDEVISE',' ' + V_PGI.SymbolePivot);
    end
    else begin
      RDev.Code := Devise.Value;
      GetInfosDevise(RDev);
      SetControlText('SYMBOLEDEVISE',' ' + RDev.Symbole);
    end;
  end;

{b fb 05/06/2006 FQ18023}
  if (OnTiers) then begin
    if (GetControlText('CPTAUX')='') then
      SetControlText('CPTAUX',GetMinMaxCompte('TIERS', 'MIN(T_AUXILIAIRE)','0'));
    if (GetControlText('CPTAUX_')='') then
      SetControlText('CPTAUX_',GetMinMaxCompte('TIERS', 'MAX(T_AUXILIAIRE)','ZZZZZZZZZZZZZZZZZ'))
   end
  else begin
    if (GetControlText('CPTAUX')='') then
      SetControlText('CPTAUX',GetMinMaxCompte('GENERAUX', 'MIN(G_GENERAL)','0'));
    if (GetControlText('CPTAUX_')='') then
      SetControlText('CPTAUX_',GetMinMaxCompte('GENERAUX', 'MAX(G_GENERAL)','ZZZZZZZZZZZZZZZZZ'));
    end;
{e fb 05/06/2006 FQ18023}

  if (GetControlText('CPTGEN')='') then
    SetControlText('CPTGEN',GetMinMaxCompte('GENERAUX', 'MIN(G_GENERAL)','0'));
  if (GetControlText('CPTGEN_')='') then
    SetControlText('CPTGEN_',GetMinMaxCompte('GENERAUX', 'MAX(G_GENERAL)','ZZZZZZZZZZZZZZZZZ'));

(*  if (VH^.ExoV8.Code <> '') then
    SetControlText('DateExoV8', UsDateTime(VH^.ExoV8.Deb) )
  else
    SetControlText('DateExoV8', UsDateTime(iDate1900) ) ;*)

  // Remplir les conditions SQL
  GenererConditionSQL;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPJUSTSOLD_QR1.AuxiElipsisClick(Sender: TObject);
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


procedure TOF_CPJUSTSOLD_QR1.GenererConditionSQL;
begin
  SqlCpt;
  SqlEcr;
end;

procedure TOF_CPJUSTSOLD_QR1.SQLEcr ;
{ Construction de la requête SQL en fonction du multicritère }
var
  StSelect, StFrom, StWhere, StOrder : string;
BEGIN
  Inherited ;
  { Construction de la clause Select de la SQL }
  StSelect := '';
  StFrom := '';
  StWhere := '';
  StOrder := '';

  if (OnTiers) then
    StSelect:=', G_LIBELLE as GLIB ';
  Case AffichageMonnaie.ItemIndex of
    0 : begin
          StSelect := StSelect + ', E_DEBIT DEBIT, E_CREDIT CREDIT, E_COUVERTURE COUVERTURE ';
        end;
    1 : begin
          StSelect := StSelect + ', E_DEBITDEV DEBIT, E_CREDITDEV CREDIT, E_COUVERTUREDEV COUVERTURE ';
        end;
  end ;

  { Construction de la clause Left Join de la SQL }
  if (OnTiers) then
    StFrom := ' Left Join GENERAUX on G_GENERAL=E_GENERAL ';

  { Construction de la clause Where de la SQL }

  // Condition sur les comptes tiers
{b fb 05/06/2006 FQ18023}
  if (OnTiers) then
    stWhere := ' AND E_AUXILIAIRE>="' + GetControlText('CPTAUX') +
  					 '" AND E_AUXILIAIRE<="' + GetControlText('CPTAUX_') + '"'
  else
    stWhere := ' AND E_GENERAL>="' + GetControlText('CPTAUX') +
                                     '" AND E_GENERAL<="' + GetControlText('CPTAUX_') + '"';
{e fb 05/06/2006 FQ18023}

  //Condition sur les comptes collectifs
  if (OnTiers) then
    stWhere := stWhere + ' AND E_GENERAL>="' + GetControlText('CPTGEN') +
  					 '" AND E_GENERAL<="' + GetControlText('CPTGEN_') + '"';

  StWhere := StWhere + ' And (E_DATECOMPTABLE<="'+USDateTime(StrToDate(DatComp.Text))+'" Or E_DATEPAQUETMIN<="'+USDateTime(StrToDate(DatComp.Text))+'") ';
  if VH^.ExoV8.Code<>'' then
    StWhere := StWhere + ' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ';

  { FQ 20389 BVE 01.06.07 if TypEcr.Value='NOR4' then }
  if TypEcr.Value='NOR' then 
    StWhere := StWhere + ' And E_QUALIFPIECE="N"';
  if TypEcr.Value='NSS' then
    StWhere := StWhere + ' And (E_QUALIFPIECE="N" Or E_QUALIFPIECE="S" Or E_QUALIFPIECE="U")';
  if TypEcr.Value='PRE' then
    StWhere := StWhere + ' And E_QUALIFPIECE="P"';
  if TypEcr.Value='SSI' then
    StWhere := StWhere + ' And (E_QUALIFPIECE="S" Or E_QUALIFPIECE="U")';
  if TypEcr.Value='TOU' then
    StWhere := StWhere + ' And (E_QUALIFPIECE="N" Or E_QUALIFPIECE="S"' +
                         ' Or E_QUALIFPIECE="U" Or E_QUALIFPIECE="P")';

  if GetControlText('REFINT')<>'' then
    StWhere := StWhere + ' And E_REFINTERNE="' + GetControlText('REFINT') + '"';
  if TypLet.Value = 'LE' then
     StWhere := StWhere + ' And E_ETATLETTRAGE="TL" '
  else if TypLet.Value = 'NL' then
     StWhere := StWhere + ' And E_ETATLETTRAGE<>"TL" ';
  StWhere := StWhere + ' And E_ETATLETTRAGE<>"RI" ';

  if SorteSens.Value = 'C' then
  begin
     Case AffichageMonnaie.ItemIndex of
        0 : begin
              StWhere := StWhere + ' And E_CREDIT<>0 ';
            end;
        1 : begin
              StWhere := StWhere + ' And E_CREDITDEV<>0 ';
            end;
      end;
  end
  else if SorteSens.Value = 'D' then
  begin
     Case AffichageMonnaie.ItemIndex of
        0 : begin
               StWhere := StWhere + ' And E_DEBIT<>0 ';
            end;
        1 : begin
               StWhere := StWhere + ' And E_DEBITDEV<>0 ';
            end;
     end;
  end;

  // Condition sur l'etablissement
  if ComboEtab.itemIndex > 0 then
    stWhere := stWhere + ' AND E_ETABLISSEMENT = "' + ComboEtab.Value + '"';

  // Condition sur la devise
  if Devise.itemIndex > 0 then
    stWhere := stWhere + ' AND E_DEVISE = "' + Devise.value + '"';

  // Condition sur le type de piece
  stWhere := stWhere + ' AND E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"CLO"'
                     + ' AND E_ECRANOUVEAU<>"OAN"' ;

  { Construction de la clause Order By de la SQL }
  StOrder := StOrder + ' Order By E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE DESC, E_LETTRAGE, E_DATECOMPTABLE, E_DATEECHEANCE, E_NUMECHE, E_NUMEROPIECE ';

  ConditionSQLEcrS.Text := StSelect;
  ConditionSQLEcrF.Text := StFrom;
  ConditionSQLEcrW.Text := StWhere;
  ConditionSQLEcrO.Text := StOrder;
end;

procedure TOF_CPJUSTSOLD_QR1.SqlCpt;
var
  Lequel,Pre,StPlus2,St1 : String ;
{$IFDEF CCMP}
  QTiers : TQuery;
  szNature : String;
{$ENDIF}
  Alias, StSelect, StFrom, StWhere, StOrder : string;
begin
  StSelect := '';
  StFrom := '';
  StWhere := '';
  StOrder := '';

  if (OnTiers) then begin
    { FQ 20392 BVE 14.06.07 }
    if ChoixTabLibres.Value = 'AUX' then
       StSelect := ',T_TABLE0,T_TABLE1,T_TABLE2,T_TABLE3,T_TABLE4,T_TABLE5,T_TABLE6,T_TABLE7,T_TABLE8,T_TABLE9'
    else if ChoixTabLibres.Value = 'CLI' then
       StSelect := ',YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3,YTC_TABLELIBRETIERS4,YTC_TABLELIBRETIERS5,YTC_TABLELIBRETIERS6,YTC_TABLELIBRETIERS7,YTC_TABLELIBRETIERS8,YTC_TABLELIBRETIERS9,YTC_TABLELIBRETIERSA'
    else if ChoixTabLibres.Value = 'SCL' then
       StSelect := ',YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3';
    { END FQ 20392 }
    StSelect := StSelect + ', T_AUXILIAIRE as CODE, T_LIBELLE as LIB, T_NATUREAUXI, T_TELEPHONE as TEL, T_ISPAYEUR as ISPAYEUR, T_PAYEUR AS PAYEUR ';
{ Tables explorées par la SQL }
    StFrom := 'TIERS T ';
    StWhere := ' AND T_LETTRABLE="X" AND ';
    Lequel:='AUXILIAIRE' ;
    Pre:='T_' ;
    Alias := 'T.';
   end
  else begin 
    { FQ 20392 BVE 14.06.07 }
    if ChoixTabLibres.Value <> '' then
       StSelect := ',G_TABLE0,G_TABLE1,G_TABLE2,G_TABLE3,G_TABLE4,G_TABLE5,G_TABLE6,G_TABLE7,G_TABLE8,G_TABLE9'
    else if ChoixTabLibres.Value = 'CLI' then
       StSelect := ',YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3,YTC_TABLELIBRETIERS4,YTC_TABLELIBRETIERS5,YTC_TABLELIBRETIERS6,YTC_TABLELIBRETIERS7,YTC_TABLELIBRETIERS8,YTC_TABLELIBRETIERS9,YTC_TABLELIBRETIERSA'
    else if ChoixTabLibres.Value = 'SCL' then
       StSelect := ',YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3';
    { END FQ 20392 }
   StSelect :=  StSelect + ', G_GENERAL as CODE, G_LIBELLE as LIB, G_NATUREGENE, G_TELEPHONE as TEL, G_TVASURENCAISS as ISPAYEUR, G_TVA AS PAYEUR ';
   { Tables explorées par la SQL }
   StFrom := StFrom + ' GENERAUX G ';
   StWhere := ' AND G_LETTRABLE="X" AND ';
   Lequel:='GENERAL' ;
   Pre:='G_' ;
   Alias := 'G.';
   end;

  { Construction de la clause Where de la SQL }
  if (V_PGI.Confidentiel<>'1') then
    StWhere := StWhere + ' '+Pre+'CONFIDENTIEL<>"1" AND ';
  if Devise.Value<>'' then
  begin
     StWhere := StWhere + ' (Exists (Select E_'+Lequel+',E_DATECOMPTABLE From ECRITURE Where E_'+Lequel+'=' +Alias+Pre+Lequel;
     StWhere := StWhere + ' And E_DEVISE="'+Devise.value+'")) ';
  end
  else
  begin
     StPlus2:='';
     if (OnTiers) then
     begin
        if CptGenDeb.Text<>'' then
           StPlus2:=StPlus2 + ' And (E_GENERAL>="'+CptGenDeb.Text+'") ' ;
        if CptGenFin.Text<>'' then
           StPlus2:=StPlus2 + ' And (E_GENERAL<="'+CptGenFin.Text+'") ' ;
     end;
//   St1:=WhereSuppEdtTiers(CritEdt.QualifPiece,V_PGI.Controleur,AvecRevision.State) ;
     if TypLet.Value = 'LE' then
     begin { Que les lettrés }
        if VH^.ExoV8.Code<>'' then
        begin
           StWhere := StWhere + ' (Exists (Select E_'+Lequel+',E_ETATLETTRAGE From ECRITURE Where E_'+Lequel+'=' +Alias+Pre+Lequel;
           StWhere := StWhere + StPlus2 + ' And (E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL") And (E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'"))) ';
        end
        else
        begin
           StWhere := StWhere + ' (Exists (Select E_'+Lequel+',E_ETATLETTRAGE From ECRITURE Where E_'+Lequel+'=' +Alias+Pre+Lequel;
           StWhere := StWhere + StPlus2 + ' And (E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL"))) ';
        end;
     end
     else if TypLet.Value = 'NL' then
     begin { Que le non lettrés }
        if VH^.ExoV8.Code<>'' then
        begin
           StWhere := StWhere + ' (Exists (Select E_'+Lequel+',E_ETATLETTRAGE From ECRITURE Where E_'+Lequel+'=' +Alias+Pre+Lequel;
           StWhere := StWhere + StPlus2 + St1 + ' And (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") And (E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'") AND (E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"OAN"))) ';
        end
        else
        begin
           StWhere := StWhere + ' (Exists (Select E_'+Lequel+',E_ETATLETTRAGE From ECRITURE Where E_'+Lequel+'=' +Alias+Pre+Lequel;
           StWhere := StWhere + StPlus2 + St1 + ' And (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") And (E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"OAN"))) ';
        end;
     end
     else
     begin
        if VH^.ExoV8.Code<>'' then
        begin
           StWhere := StWhere + ' (Exists (Select E_'+Lequel+',E_DATECOMPTABLE From ECRITURE Where E_'+Lequel+'=' +Alias+Pre+Lequel;
           StWhere := StWhere + StPlus2 + St1 +' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" )) ';
        end
        else
           StWhere := StWhere + ' '+Pre+Lequel+' <> "*" ';
     end;
  end;
  if CptAuxDeb.Text<>'' then
     StWhere := StWhere + ' AND '+Pre+Lequel+'>="'+CptAuxDeb.Text+'" ';
  if CptAuxFin.Text<>'' then
     StWhere := StWhere + ' AND '+Pre+Lequel+'<="'+CptAuxFin.Text+'" ';
  if NatComp.Value<>'' then
  begin
     if (OnTiers) then
        StWhere := StWhere + ' AND T_NATUREAUXI="'+NatComp.Value+'" '
     else
        StWhere := StWhere + ' AND G_NATUREGENE="'+NatComp.Value+'" ';
  end
  else
  begin
     {$IFDEF CCMP}
     QTiers := OpenSQL('SELECT T_NATUREAUXI FROM TIERS WHERE T_AUXILIAIRE="'+CptAuxDeb.Text+'"',True);
     szNature := QTiers.Fields[0].AsString;
     Ferme(QTiers);

     if (szNature='TID') then
        StWhere := StWhere + ' AND (G_NATUREGENE="TID") ';
     if (szNature='CLI') then
        StWhere := StWhere + ' AND ((T_NATUREAUXI="CLI") OR (T_NATUREAUXI="AUD")) ';

     if (szNature='TIC') then
        StWhere := StWhere + ' AND (G_NATUREGENE="TIC") ';
     if (szNature='FOU') then
        StWhere := StWhere + ' AND ((T_NATUREAUXI="FOU") OR (T_NATUREAUXI="AUC")) ';
     {$ELSE}
     if not (OnTiers) then
        StWhere := StWhere + ' AND (G_NATUREGENE="TID" OR G_NATUREGENE="TIC")'
     else
        StWhere := StWhere + ' AND T_NATUREAUXI<>"NCP" ';
     {$ENDIF}
  end;

  if (OnTiers) then begin
    if ExCompSold.Checked then
      StWhere := StWhere + ' AND (T_TOTALDEBIT<>T_TOTALCREDIT) ';
   end
  else begin
    if ExCompSold.Checked then
      StWhere := StWhere + ' AND (G_TOTALDEBIT<>G_TOTALCREDIT) ';
    end;

  { Construction de la clause Order By de la SQL }
  Case TriLib.Checked of
    false : if (OnTiers) then
              StOrder := ' T_NATUREAUXI, T_AUXILIAIRE '
            else
              StOrder := ' G_NATUREGENE, G_GENERAL ';
    true : if (OnTiers) then
              StOrder := ' T_NATUREAUXI, T_LIBELLE '
           else
              StOrder := ' G_NATUREGENE, G_LIBELLE ';
    end;

  ConditionSQLCptS.Text := StSelect;
  ConditionSQLCptF.Text := StFrom;
  ConditionSQLCptW.Text := StWhere;
  ConditionSQLCptO.Text := StOrder;
end ;

{b fb 05/06/2006 FQ18023}
function TOF_CPJUSTSOLD_QR1.GetWhereGenTie(stNatureGene: String): String;
begin
  result := ' WHERE G_NATUREGENE="TID" OR G_NATUREGENE="TIC" ';
  if stNatureGene='TID' then
    result := ' WHERE G_NATUREGENE="TID" '
  else if stNatureGene='TIC' then
    result := ' WHERE G_NATUREGENE="TIC" ';
end;
{e fb 05/06/2006 FQ18023}

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 02/05/2007
Modifié le ... : 02/05/2007
Description .. : Permet de completer automatiquement la zone auxiliaire
Suite ........ : FQ20167
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTSOLD_QR1.CompteOnExit(Sender: TObject);
begin
  inherited; 
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if TestJoker(CptAuxDeb.Text) then Exit;

  // Complétion auto du numéro de compte si possible
  if ExisteSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = "' + BourreEtLess(THEdit(Sender).Text, fbAux) + '"') then
  begin
     if Length(THEdit(Sender).Text) < VH^.Cpta[fbAux].Lg then
        THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbAux);
  end
  else
     THEdit(Sender).ElipsisClick(nil);
end;

Initialization
  registerclasses ( [ TOF_CPJUSTSOLD_QR1 ] ) ;
end.







