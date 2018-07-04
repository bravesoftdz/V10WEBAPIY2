unit UTofAFPlanning_Mul;

interface

uses  StdCtrls,Controls,Classes,M3FP,HTB97,AFPlanning, AFPlanningCst, hqry,
{$IFDEF EAGLCLIENT}
      eMul,utob,Maineagl,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,mul, FE_main,
{$ENDIF}
{$IFDEF BTP}
	  CalcOleGenericBTP,
{$ENDIF}
      forms,sysutils,HDB,
      ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ConfidentAffaire,
      AffaireUtil,AffaireRegroupeUtil,Dicobtp,SaisUtil,EntGC,
      utofAfBaseCodeAffaire,utilpgi,AglInit,UtilGc,TraducAffaire,
      UtilTaches;
                           
Type
     TOF_AFPLANNING_MUL = Class (TOF_AFBASECODEAFFAIRE)
        procedure OnArgument(stArgument : String ); override;
        procedure OnUpdate; override;
        procedure OnLoad; override;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
        //procedure TraiteRetour;
        procedure BOuvrirOnClick(Sender : TObject);

     private
{$IFDEF EAGLCLIENT}
       fListe : THGrid;
{$ELSE}
       fListe : THDBGrid;
{$ENDIF}
     End;

const
	TexteMsgTache: array[1..2] of string 	= (
          {1}        'Cette affaire ne contient aucune tâche en mode de saisie quantité',
          {2}        'Cette affaire ne contient aucune tâche en mode de saisie montant');

  //function AFLanceFicheAFPlanning(Lequel, Argument : String) : String;
  function AFLanceFicheAFPlanning_Mul(Lequel, Argument : String) : String;

implementation           

Function AFLanceFicheAFPlanning_Mul(Lequel, Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFPLANNING_MUL','', Lequel,Argument);
end;
                                
procedure TOF_AFPLANNING_MUL.OnArgument(stArgument : String);
var
  vStTmp    : String;
  vStChamp  : String;
  vStValeur, Aff0,Aff1,Aff2,Aff3,Avenant : String;

Begin

  Inherited;

  // traitement des arguments
  vStTmp:= (Trim(ReadTokenSt(stArgument)));
  While (vStTmp <>'') do
    Begin
      DecodeArgument(vStTmp, vStChamp, vStValeur);
      if vStChamp = 'PLANNING1' then
        Begin
          TFMUL(ecran).DBListe := 'AFMULPLANNING';
          Ecran.Caption := TraduitGA('Tâches par Affaire/Fonction/Ressource');
          UpdateCaption(Ecran);
          setControlText('SELECTION','PLANNING1');
        End                            

      else if vStChamp = 'PLANNING3' then
        Begin
          TFMUL(ecran).DBListe := 'AFMULPLANNING';
          Ecran.Caption := TraduitGA('Ressources par Affaire/Tâche');
          UpdateCaption(Ecran);
          setControlText('SELECTION','PLANNING3');
        End
      Else if vStChamp = 'AFFAIRE' then   //mcd 04/2003 pour appel depuis TB
        begin
        setcontrolText ('AFF_AFFAIRE',vStValeur);
        {$IFDEF BTP}
		BTPCodeAffaireDecoupe (VstValeur,aff0,aff1,Aff2,Aff3,avenant,taconsult,false);
        {$ELSE}
        CodeAffaireDecoupe (VstValeur,aff0,aff1,Aff2,Aff3,avenant,taconsult,false);
        {$ENDIF}
        SetControlText ('AFF_AFFAIRE1',Aff1);
        SetControlText ('AFF_AFFAIRE2',Aff2);
        SetControlText ('AFF_AFFAIRE3',Aff3);
        SetControlText ('AFF_AVENANT',AVenant);
        end
      Else if vStChamp = 'TIERS' then setcontrolText ('AFF_TIERS',vStValeur)

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

   TFMul(Ecran).BOuvrir.OnClick := BOuvrirOnClick ;
   fListe := TFMul(Ecran).FListe;

   // temporaire
   // le dev n'est pas terminé
   SetControlVisible('CONGES', False);
  {$IFDEF CCS3}
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  if (getcontrol('PSTAT') <> Nil) then SetControlVisible ('PSTAT', False);
  {$ENDIF}
End;
      
procedure TOF_AFPLANNING_MUL.OnUpdate;
Begin

  // Gestion repositionnement auto sur l'affaire en cours si sortie rapide / bug par prg ( ou Eagl)
  inherited;

  if Not (VH_GC.CleAffaire.Co2Visible) then SetControlText('AFF_AFFAIRE2','');
  if Not (VH_GC.CleAffaire.Co3Visible) then SetControlText('AFF_AFFAIRE3','');

End;

{procedure TOF_AFPLANNING_MUL.TraiteRetour;
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

procedure TOF_AFPLANNING_MUL.OnLoad;
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
end;

procedure TOF_AFPLANNING_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
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

//******************************************************************************
//************* Fonctions appellées depuis le Scripts ***************************
//******************************************************************************
{procedure MajMulAfPlanning( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (LaTof is TOF_AFPLANNING_MUL) then TOF_AFPLANNING_MUL(LaTof).TraiteRetour else exit;
end;
}

procedure TOF_AFPLANNING_MUL.BOuvrirOnClick(Sender : TObject);
var
  i                 : Integer;
  vStWhere          : String;
  vStEtats          : String;
  vStEtat           : String;
  vStWhereEtatLigne : String;
  vStConges         : String;

begin

  vStWhere := Copy(RecupWhereCritere(TPageControl(GetControl('PAGES'))),6,length(RecupWhereCritere(TPageControl(GetControl('PAGES')))) );

  // ajout du critere etat des lignes de planning
  if GetControlText('ETATLIGNE') <> '' then
  begin
    i := 0;
    vStEtats := GetControlText('ETATLIGNE');
    While vStEtats <> '' do
    begin
      vStEtat := ReadTokenSt(vStEtats);
      if i = 0 then
        vStWhereEtatLigne := ' AND (APL_ETATLIGNE = "' + vStEtat + '"'
      else
        vStWhereEtatLigne := vStWhereEtatLigne + ' OR APL_ETATLIGNE = "' + vStEtat + '"';
      i := i + 1;
    end;
    vStWhereEtatLigne := vStWhereEtatLigne + ') ';
  end;

  // on ajoute aux criteres la liste des affaires sélectionnées
  try
    if (not fListe.AllSelected) and (fListe.nbSelected <> 0) then
      begin
        for i := 0 to fListe.nbSelected -1 do
          begin
            fListe.GotoLeBookMark(i);
            if (i = 0) and (vStWhere <> '') then vStWhere := vStWhere + ' AND ( '
            else if (i <> 0) and (vStWhere <> '') then vStWhere := vStWhere + ' OR ';
            vStWhere := vStWhere + 'AFF_AFFAIRE = "' + TFMul(Ecran).Q.Findfield('AFF_AFFAIRE').AsSTring + '"';
          end;
        if vStWhere <> '' then vStWhere := vStWhere + ')';
      end;
   except
   on E:Exception do
    begin
     MessageAlerte('Erreur lors du chargement du planning' +#10#13#10#13 + E.message ) ;
    end; // on
   end; // try

  if TCheckBox(GetControl('CONGES')).Checked then
    vStConges := 'X'
  else
    vStConges := '-';

  {if (GetControlText('SELECTION') = 'PLANNING1') or
     (GetControlText('SELECTION') = 'PLANNING3') then
    ExecPlanning(GetControlText('NUMPLANNING'), '01/01/1900', vStWhere, vStWhereEtatLigne, vStConges);}
  
end;

Initialization
  registerclasses([TOF_AFPLANNING_MUL]);
//  RegisterAglProc('MajMulAfPlanning', TRUE , 0, MajMulAfPlanning);
end.
