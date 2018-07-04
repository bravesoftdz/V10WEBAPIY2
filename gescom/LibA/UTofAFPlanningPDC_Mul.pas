unit UTofAFPlanningPDC_Mul;

interface

uses  StdCtrls,Controls,Classes,M3FP,HTB97,AFPlanning, AFPlanningCst, hqry,
{$IFDEF EAGLCLIENT}
      eMul,utob, Maineagl,
{$ELSE}
      db,dbTables,DBGrids,mul, FE_main,
{$ENDIF}
      forms,sysutils,HDB,CalcOleGenericAff,
      ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ConfidentAffaire,
      AffaireUtil,AffaireRegroupeUtil,DicoAF,SaisUtil,EntGC,
      utofAfBaseCodeAffaire,utilpgi,AglInit,UtilGc,TraducAffaire,
      UtilTaches, UTofAfPlanCharge, UTofAfTaches;

Type
     TOF_AFPLANNINGPDC_MUL = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument(stArgument : String ); override;
        procedure OnUpdate; override;
        procedure OnLoad; override;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
        //procedure TraiteRetour;
        function ExisteTache(pStAffaire : String) : Boolean;
        procedure FListe_OnDblClick(Sender : TObject);
     End;

  Function AFLanceFicheAFPlanningPDC_mul(Lequel, Argument : String) : String;

const
	TexteMsgTache: array[1..2] of string 	= (
          {1}        'Cette affaire ne contient aucune tâche en mode de saisie quantité',
          {2}        'Cette affaire ne contient aucune tâche en mode de saisie montant');

implementation

Function AFLanceFicheAFPlanningPDC_mul(Lequel, Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFPLANNINGPDC_MUL','', Lequel, Argument);
end;

procedure TOF_AFPLANNINGPDC_MUL.OnArgument(stArgument : String);
var
  vStTmp    : String;
  vStChamp  : String;
  vStValeur : String;

Begin

  Inherited;

  // traitement des arguments
  vStTmp:= (Trim(ReadTokenSt(stArgument)));
  While (vStTmp <>'') do
    Begin
      DecodeArgument(vStTmp, vStChamp, vStValeur);
      If vStChamp = 'PLANCHARGEQTE' then
        Begin
          TFMUL(ecran).DBListe := 'AFMULPLANCHARGE';
          Ecran.Caption := TraduitGA('Plan de Charge en quantité');
          UpdateCaption(Ecran);
          setControlText('SELECTION','PLANCHARGEQTE');
        End
      else If vStChamp = 'PLANCHARGEMNT' then
        Begin
          TFMUL(ecran).DBListe := 'AFMULPLANCHARGE';
          Ecran.Caption := TraduitGA('Plan de Charge en montant');
          UpdateCaption(Ecran);
          setControlText('SELECTION','PLANCHARGEMNT');
        End

      Else if vStChamp = 'TACHES' then
        Begin
          TFMUL(ecran).DBListe := 'AFMULAFFTACHE';
          Ecran.Caption := TraduitGA('Accès par affaire');
          UpdateCaption(Ecran);
          setControlText('SELECTION','TACHES');
        end

      Else if vStChamp = 'NUMPLANNING' then
        setControlText('NUMPLANNING',vStValeur);

      vStTmp:=(Trim(ReadTokenSt(stArgument)));
    End;

  // on ne peut pas modifier les affaires
  SetControlVisible ('BInsert',False);
  SetControlVisible ('BInsert1',False);
  SetControlVisible('bAssistantCreation',False);

  // et les sous affaires
  if not(GereSousAffaire) then
    begin
      SetcontrolVisible('TAFF_AFFAIREREF',False);
      SetcontrolVisible('AFFAIREREF1',False); SetcontrolVisible('AFFAIREREF2',False);
      SetcontrolVisible('AFFAIREREF3',False); SetcontrolVisible('AFFAIREREF4',False);
      SetcontrolVisible('AFF_ISAFFAIREREF',False);
      SetcontrolVisible('BSELECTAFF2',False);
    end;

  {$IFDEF EAGLCLIENT}
    TraduitAFLibGridSt(TFMul(Ecran).FListe);
  {$ELSE}
    TraduitAFLibGridDB(TFMul(Ecran).FListe);
  {$ENDIF}

   TFMul(Ecran).FListe.OnDblClick := FListe_OnDblClick;

   // C.B 15/10/03
   // Forcer PROP unchecked -> sans les propositions
   TCheckbox(GetControl('PROP')).State := cbUnchecked;
  {$IFDEF CCS3}
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  if (getcontrol('PSTAT') <> Nil) then SetControlVisible ('PSTAT', False);
  {$ENDIF}
End;                                          

procedure TOF_AFPLANNINGPDC_MUL.OnUpdate;
Begin

  // Gestion repositionnement auto sur l'affaire en cours si sortie rapide / bug par prg ( ou Eagl)
  inherited;

  if Not (VH_GC.CleAffaire.Co2Visible) then SetControlText('AFF_AFFAIRE2','');
  if Not (VH_GC.CleAffaire.Co3Visible) then SetControlText('AFF_AFFAIRE3','');

End;

{procedure TOF_AFPLANNINGPDC_MUL.TraiteRetour;
Var stRetour  : string;
    Champ     : string;

begin

  stRetour := GetControlText ('RETOUR');
  if stRetour <> '' then
    begin
      // On récupére dans l'ordre le client + Affaire 1,2,3
      Champ:=(Trim(ReadTokenSt(stRetour))); if Champ <> '' then SetControltext('AFF_TIERS',Champ);
      Champ:=(Trim(ReadTokenSt(stRetour))); if Champ <> '' then SetControltext('AFF_AFFAIRE1',Champ);
      Champ:=(Trim(ReadTokenSt(stRetour))); if (Champ <> '') And (VH_GC.CleAffaire.Co2Visible) then SetControltext('AFF_AFFAIRE2',Champ);
      Champ:=(Trim(ReadTokenSt(stRetour))); if (Champ <> '') And (VH_GC.CleAffaire.Co3Visible) then SetControltext('AFF_AFFAIRE3',Champ);
    end;
  SetControlText('RETOUR', '');

end;
}

procedure TOF_AFPLANNINGPDC_MUL.OnLoad;
Var Affaire,Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant : string;

begin

  inherited;
  if GereSousAffaire then
    begin
      if (GetControlText('AFFAIREREF1')='') then SetControlText('AFF_AFFAIREREF','')
      else
        begin
          Affaire0 := 'A';
          Affaire1 := GetControlText('AFFAIREREF1'); Affaire2 := GetControlText('AFFAIREREF2');
          Affaire3 := GetControlText('AFFAIREREF3'); Avenant := GetControlText('AFFAIREREF4');
          Affaire:=CodeAffaireRegroupe(Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant, taModif, false,false,false);
          if not ExisteAffaire(Affaire,'') then SetControlText('AFF_AFFAIREREF',Trim(Copy(Affaire,1,15)));
        end;
    end;
//  V_PGI.NbOpenQuery := V_PGI.NbOpenQuery - 1; 

end;

procedure TOF_AFPLANNINGPDC_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('AFF_AFFAIRE'));
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1')); Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3')); Aff4:=THEdit(GetControl('AFF_AVENANT'));
  Tiers:=THEdit(GetControl('AFF_TIERS'));

  // affaire de référence pour recherche
  Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
  Aff1_:=THEdit(GetControl('AFFAIREREF1'));
  Aff2_:=THEdit(GetControl('AFFAIREREF2'));Aff3_:=THEdit(GetControl('AFFAIREREF3'));
  Aff4_:=THEdit(GetControl('AFFAIREREF4'));
end;

function TOF_AFPLANNINGPDC_MUL.ExisteTache(pStAffaire : String) : Boolean;
var
  vQr               : TQuery;
  S                 : String;

begin

  Result := True;

  S := 'select ATA_LIBELLETACHE1, ATA_QTEINITIALE, ATA_FONCTION, ';
  S := S + ' ATA_QTEAPLANIFIER, ATA_NUMEROTACHE, ATA_ARTICLE ';
  S := S + ' FROM AFFAIRE, TACHE ';
  S := S + ' where AFF_AFFAIRE = "' + pStAffaire + '"';

  if GetControlText('SELECTION') = 'PLANCHARGEMNT' then
    S := S + ' and ATA_MODESAISIEPDC = "MPR" '
  else
    S := S + ' and ATA_MODESAISIEPDC = "QUA" ';

  S := S + ' and AFF_AFFAIRE = ATA_AFFAIRE ';
  S := S + ' ORDER BY ATA_NUMEROTACHE';

  vQr := nil;
  Try
    vQr := OpenSql(S, True);
    if vQr.Eof then
      begin
        if GetControlText('SELECTION') = 'PLANCHARGEQTE' then
          begin
            PGIBoxAF (TexteMsgTache[1],'');
            Result := false;
          end
        else if GetControlText('SELECTION') = 'PLANCHARGEMNT' then
          begin
            PGIBoxAF (TexteMsgTache[2],'');
            Result := false;
          end;
      end;
  Finally
    if vQr <> Nil then Ferme(vQr);
  End;
end;

procedure TOF_AFPLANNINGPDC_MUL.FListe_OnDblClick(Sender : TObject);
begin
  if GetControlText('SELECTION') = 'PLANCHARGEQTE' then
    begin
      if ExisteTache(GetField('AFF_AFFAIRE')) then
        AFLanceFicheAFPlanCharge(GetField('AFF_AFFAIRE'),
                                 'AFFAIRE:' + GetField('AFF_AFFAIRE') +
                                 ';LIBAFFAIRE:' + CodeAffaireAffiche(GetField('AFF_AFFAIRE')) + '  ' + GetField('AFF_LIBELLE') +
                                 ';PLANCHARGEQTE; MONOFICHE');
    end

  else if GetControlText('SELECTION') = 'PLANCHARGEMNT' then
    begin
      if ExisteTache(GetField('AFF_AFFAIRE')) then
        AFLanceFicheAFPlanCharge(GetField('AFF_AFFAIRE'),
                                 'AFFAIRE:' + GetField('AFF_AFFAIRE') +
                                 ';LIBAFFAIRE:' + CodeAffaireAffiche(GetField('AFF_AFFAIRE')) + '  ' + GetField('AFF_LIBELLE') +
                                 ';PLANCHARGEMNT; MONOFICHE');
    end
 
  else if GetControlText('SELECTION') = 'TACHES' then
    AFLanceFicheAFTaches('ATA_AFFAIRE:' + GetField('AFF_AFFAIRE') + ';ACTION=MODIFICATION');
end;

//******************************************************************************
//************* Fonctions appellées depuis le Scripts ***************************
//******************************************************************************
{procedure MajMulAfPlanningPdc( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (LaTof is TOF_AFPLANNINGPDC_MUL) then TOF_AFPLANNINGPDC_MUL(LaTof).TraiteRetour else exit;
end;
}

Initialization
  registerclasses([TOF_AFPLANNINGPDC_MUL]);
//  RegisterAglProc('MajMulAfPlanningPdc', TRUE , 0, MajMulAfPlanningPdc);
end.
