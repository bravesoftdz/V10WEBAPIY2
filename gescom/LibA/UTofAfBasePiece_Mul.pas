unit UTofAfBasePiece_Mul;
                                                         
// mcd 09/10/2002 tout changé pour le traitement de la zone naturepièce qui
// peut maintenant être en multivalcombobox ou en valcombobox (comme avant)

interface
uses
{$IFDEF EAGLCLIENT}
     eMul,
{$ELSE}
     Mul,  db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
      ParamSoc,
      CalcOleGenericBTP,
{$ENDIF}
      StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,UTOFAFBASEMODELEEDIT,
      HCtrls,HEnt1,HMsgBox,UTOF, entgc,AffaireUtil,factutil,Dicobtp,SaisUtil,
      HTB97, M3Fp, UTOB, UtilGc;

Type
     TOF_AFBASEPIECE_MUL = Class (TOF_AFBASEMODELEEDIT)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnLoad ; override ;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
        procedure NomsChampsNature(ztable : string; var  App:THEdit;
         var TApp:THLabel; var  Viv:TCheckBox); 
        procedure ChargeTobComplete(TobComplete: Tob);
        procedure AlimPieceVivante;
     public
         NatPiece       :   ThValCombobox;
         NatPieceMul    :   ThMultiValCombobox;
         EditApp        :   THEdit;
         EditTApp       :   THLabel;
         Vivante        :   TCheckBox;
         Nochange       :   boolean;
         sNature        :   string;
     END ;

implementation

uses uTOFComm;

procedure TOF_AFBASEPIECE_MUL.OnArgument(stArgument : String );
var st,critere : string;
    PlusCombo:string;
    DiffTop:integer;
    Champ, valeur,ztable ,ZnatAuxi,sNatureCoche,ZdateD,ZdateF, part0,part1, part2, PArt3,Avenant : String;
    x : integer;
    queprovisoires : boolean;
    sStatutLiberte : string;
Begin
fMulDeRecherche := true;
Inherited;
//
// Recup des critères
//  NOCHANGE_NATUREPIECE  :  pour interdire la saisie de  naturepiece
//  TABLE:GP              :  Pour savoir si on est au niveau piece GP ou ligne GL
//  NATURE:FPR            :  Nature qui va initialiser le mul
//  DUPPLI:FPR            :  Nature de la piece à dupliquer
//  NUMDEB : et NUMFIN :  :  Borne debut et fin (cf utofafpiecepro)
//  ACTION=CREATION       :  Mode consultation ou modification
// NATUREAUXI = CLI ou FOU : depuis appel fiche client/fournisseure
sNatureCoche :='';
critere :=''; St:=stArgument;
Nochange:= false; queprovisoires := false;
snature:=''; // gm 070502
ZdateD:='';
ZdateF:='';
sStatutLiberte := 'AFF';
While St<>'' do
   BEGIN
   Critere:=(Trim(ReadTokenSt(st)));
   if (Critere ='NOCHANGE_NATUREPIECE') then
      begin
      Nochange:= true;
      if (Ecran).name = 'AFPIECE_MUL' then TFmul(Ecran).FiltreDisabled:=TRue; // mcd 13/12/02 car même fiche pour  appel
      end
   else if (Critere ='QUEPROVISOIRES') then
      queprovisoires:= true
   else if (Critere ='NOVISIBLE_AFFAIRE') then
      begin
      SetControlVisible('TGP_AFFAIRE',false);
      SetControlVisible('BSELECTAFF1',false);
      SetControlVisible('BEFFACEAFF1',false);
      SetControlVisible('AFF_AFFAIRE1',false);
      SetControlVisible('AFF_AFFAIRE2',false);
      SetControlVisible('AFF_AFFAIRE3',false);
      SetControlVisible('AFF_AVENANT',false);
      SetControlVisible('TGP_TIERS',false);
      SetControlVisible('GP_TIERS',false);
      SetControlVisible('BPOPMENU',false);
      DiffTop := GetControl('GP_REFINTERNE').Top - GetControl('AFF_AFFAIRE1').Top;
      GetControl('GP_REFINTERNE').Top := GetControl('GP_REFINTERNE').Top-DiffTop;
      GetControl('TGP_REFINTERNE').Top := GetControl('TGP_REFINTERNE').Top-DiffTop;
      GetControl('TGP_NUMPIECE').Top := GetControl('TGP_NUMPIECE').Top-DiffTop;
      GetControl('TGP_NUMPIECE__').Top := GetControl('TGP_NUMPIECE__').Top-DiffTop;
      GetControl('GP_NUMERO').Top := GetControl('GP_NUMERO').Top-DiffTop;
      GetControl('GP_NUMERO_').Top := GetControl('GP_NUMERO_').Top-DiffTop;
      end
   else  if Critere <>'' then
        begin
        X:=pos(':',Critere);
        if x=0 then  X:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if champ = 'STATUT' then sStatutLiberte := Valeur else
        if Champ = 'NATURE' then sNature:= Valeur
        else if Champ = 'TABLE' then Ztable := Valeur
        else if Champ = 'NATUREAUXI' then Znatauxi := Valeur
        else if Champ = 'DATEDEB' then ZdateD:=Valeur
        else if Champ = 'AFFAIRE' then
           begin   //mcd 10/06/02
           if ThEdit(GetControl('GL_AFFAIRE'))<>Nil then
             begin
             SetControlText('GL_AFFAIRE',Valeur);
			{$IFDEF BTP}
             BTPCodeAffaireDecoupe(GetControlText('GL_AFFAIRE'),Part0,Part1,Part2,Part3,Avenant,taModif,false);
            {$ELSE}
             CodeAffaireDecoupe(GetControlText('GL_AFFAIRE'),Part0,Part1,Part2,Part3,Avenant,taModif,false);
            {$ENDIF}
             SetControlText('GL_AFFAIRE1',Part1); SetControlText('GL_AFFAIRE2',Part2);
             SetControlText('GL_AFFAIRE3',Part3); SetControlText('GL_AVENANT',avenant);
             end;
           end
        else if Champ = 'DATEFIN' then ZdateF:=Valeur
        end;
   End;


if (ztable = '') then ztable := 'GP';

If ZdateD <>'' then
  begin    //mcd 04/2003
  if Ztable ='GL' then  begin
      SetControlText('GL_DATEPIECE',Zdated);
      If ZdateF<>'' then  SetControlText('GL_DATEPIECE_',Zdatef);
      end
   else begin
     SetControlText('GP_DATEPIECE',Zdated) ;
     If ZdateF<>'' then  SetControlText('GP_DATEPIECE_',Zdatef);
     end;
  end;
NomsChampsNature(ztable, EditApp, EditTApp,Vivante);
{$IFNDEF BTP}
if ((NatPiece<>nil) and (Nochange = true ))then  NatPiece.Enabled := false;
if ((NatPieceMul<>nil) and (Nochange = true ))then  NatPieceMul.Enabled := false;
{$ENDIF}


//
// Gestion de la notion XX_VIVANTE si facture PROVISOIRE ou autre pièce
//

{$IFDEF BTP}
if (sStatutLiberte = 'INT') or (sStatutLiberte = 'APP') or (sStatutLiberte = 'GRP') then
begin
  if (sNature <> '') then   //gm 070502
  Begin
   if (sNature = 'FPR') or (sNature = 'APR') or (sNature = 'FRE')  then
   begin
       if vivante <>Nil then Vivante.checked := true
   end else
   begin
    if (sNature = 'FRE') then   SetControlProperty('GP_VIVANTE','State',cbgrayed)
       else if vivante <>Nil then Vivante.checked := false;  // mcd 12/08/02 ajout test Nil
    end;
  End;
end else
begin
  SetControlProperty('GP_VIVANTE','State',cbgrayed);
  SetControlVisible('GP_VIVANTE',false);
end;
{$ELSE}
if (sNature <> '') then   //gm 070502
Begin
 if (sNature = 'FPR') or (sNature = 'APR') or (sNature = 'FRE')  then
     if vivante <>Nil then Vivante.checked := true
 else
  if (sNature = 'FRE') then   SetControlProperty('GP_VIVANTE','State',cbgrayed)
     else if vivante <>Nil then Vivante.checked := false;  // mcd 12/08/02 ajout test Nil
End;
{$ENDIF}

if (ZnatAuxi<>'FOU') and        //gm 27/12/02
	((Ecran.name='AFREGRLIGNE_MUL') or (Ecran.name='AFPIECECOURS_TIE')) then

   begin  // si appel depuis clt, on ajoute nature FRE pour voir les pièces
          // idem que source AglInitAff, modif du 13/02/02
   sNature:='FR';
   end;
//
// Gestion des natures de pièces exclues
//
PlusCombo:='';
              // mcd tout revu pour passer par les natures gérée
if ZnatAuxi='FOU' then begin
      // cas appel depuis fiche fournisseur, on affiche que les
      // nature du module achat
    if (ztable ='GL') then
    	SetControlProperty( 'GL_NATUREPIECEG','Datatype','GCNATUREPIECEGACH')
    else
    	SetControlProperty( 'GP_NATUREPIECEG','Datatype','GCNATUREPIECEGACH');
     PlusCOmbo := Pluscombo + AfPlusNatureAchat;
     If CtxScot in V_Pgi.Pgicontexte then sNatureCoche := 'FF;AF;'
     else  sNatureCoche := 'CF;BLF;FF;AF;';
     if sNature='' then sNature:='FF';
     SetControlVisible ('TAFF_AFFAIRE' , False);
     SetControlVisible ('BSELECTAFF1' , False);
     SetControlVisible ('AFF_AFFAIRE1' , False);
     SetControlVisible ('AFF_AFFAIRE2' , False);
     SetControlVisible ('AFF_AFFAIRE3' , False);
     SetControlVisible ('TGP_APPORTEUR' , False);
     SetControlVisible ('GP_APPORTEUR' , False);
     If GetControl('TGP_TIERSFACTURE') <>Nil then SetControlText('TGP_TIERSFACTURE' , 'Fournisseur facturé');
     If GetControl('TGP_TIERSPAYEUR') <>Nil then SetControlText('TGP_TIERSPAYEUR' , 'Fournisseur payeur');
     end
else begin
     if (sNature = '') then sNature := 'FAC';

     PlusCombo:=PlusCombo+' AND (';

    if (ctxGCAff in v_PGI.PGIContexte) or  (ctxtempo in v_PGI.PGIContexte)then
    begin
      if (sNature = 'FR') then   sNature:='';   /// inutile pour accés via gc
      if  not queprovisoires then  // gm 11/04/02
        begin
        PlusCombo:=PlusCombo+' ((GPP_NATUREPIECEG<>"'+ VH_GC.AFNatAffaire+ '")';
        PlusCombo:=PlusCombo+' AND (GPP_NATUREPIECEG<>"'+ VH_GC.AFNatProposition+ '"))';
        sNatureCoche := 'FAC;FPR;FRE;AVC;APR';
        end
      else
        begin
         PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FPR")';
         PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="APR")';
         sNatureCoche := 'FPR;APR';
        End;
    end
    else
    begin
{$IFDEF BTP}
      if not noChange then
      begin
    		if (sStatutLiberte = 'INT') OR (sStatutLiberte = 'APP') OR (sStatutLiberte = 'GRP') then
        begin
          if not queprovisoires then
              begin
              PlusCombo:=PlusCombo+' (GPP_NATUREPIECEG="AVC")';
              PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FAC")';
              PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FRE") OR ';   // gm 20/09/02
              end;
          PlusCombo:=PlusCombo+' (GPP_NATUREPIECEG="FPR")';
          PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="APR")';
        end
        else
        begin
          PlusCombo:=PlusCombo+' (GPP_NATUREPIECEG="ABT")';
				  PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="AFF")';
          PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="DBT")';
          PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FBT")';
          PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="PBT")';
				  PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FPR")';
          PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FAC")';
          PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="AVC")';
          PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="DAP")';
          PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="ETU")';
          sNatureCoche := 'ABT;AFF;DBT;DAP;FBT;PBT;FPR;FAC;AVC;ETU';
          if GetParamSoc('SO_BTLIVRVISIBLE') then
          begin
            PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="BLC")';
            PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="LBT")';
            sNatureCoche := sNatureCoche + ';BLC;LBT';
          end;
        end;
		end
    else
    begin
      PlusCombo := PlusCombo + ' (GPP_NATUREPIECEG="'+sNature+'")';
      NatPiece.Value := sNature;
    end;
{$ELSE}
      if not queprovisoires then
      begin
        PlusCombo:=PlusCombo+' (GPP_NATUREPIECEG="AVC")';
        PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FAC")';
        PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FRE")';   // gm 20/09/02
      end;
      PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FPR")';
      PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="APR")';
{$ENDIF}
      if (sNature='FSI') then PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FSI")';
      if (sNature='FRE') then PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FRE")';
      if (sNature='FR')  then
      begin // mcd 13/02/02 pour , depuis les missions avoir FRE en plus fac et avc, sans se positioner sur cette valeur
        PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="FRE")';
        sNature:='';
      end;
      if ctxGCAFF in V_PGI.PGIContexte then
      begin
        PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="PRE")';
        PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="PRE")';
        PlusCombo:=PlusCombo+' OR (GPP_NATUREPIECEG="PRO")';
      end;
    end;
    PlusCombo:=PlusCombo+' )';
  end ;

  if (PlusCombo<>'') and (NatPiece <>Nil)    then NatPiece.Plus := PlusCombo;
  if (PlusCombo<>'') and (NatPieceMul <>Nil) then
  begin
    NatPieceMul.Plus := PlusCombo;
    // mcd 09/12/02dans le cas de multicombobox, il faut remplir le .text si on veut que le select all focntionne
    if (NatPieceMul.text = '') then
    // Si on a un champ Etat Affaire et qu'il n'est pas préinitialisé, il faut adapter la sélection du mul
    // en mettant toutes les valeurs présentes dans la liste dans la propriété .Text
      NatPieceMul.Text := StringReplace(NatPieceMul.Values.CommaText, ',', ';', [rfReplaceAll]);
  end;

  // Alimenter la valeur APRES la propriéré plus, car cette dernière recharge
  // et réinitailise

  if (NatPiece <>Nil)  then Natpiece.value := sNature;

  {$IFDEF BTP}
  if ((NatPiece<>nil) and (Nochange = true ))     then  NatPiece.Enabled := false;
  if ((NatPieceMul<>nil) and (Nochange = true ))  then  NatPieceMul.Enabled := false;
  {$ENDIF}

  // if (NatPieceMul <>Nil)  and (sNature<>'')  then NatpieceMul.Text := sNature;
	if (NatPieceMul <> Nil)  then
  Begin
    if (sNature <> '')  then
    	NatpieceMul.Text := sNature
    else
			if sNatureCoche<> '' then 	NatpieceMul.text  := sNatureCoche
  End;

  if (NatPieceMul <> Nil)  and (ZnatAuxi='FOU')  and  not(ctxscot in v_PGI.PGIContexte) then
		 NatpieceMul.text  := sNatureCoche;

  // si apporteur non géré, on met la zone en invisible
  if Not(VH_GC.AFGestionCom) then
  begin
     if (EditApp<>nil) then EditApp.visible:=false;
     if (EditTApp<>nil) then EditTApp.visible:=false;
  end;
// mcd 10/06/03 si appel depuis client ou mission, il ne faut pas modifier client, même sur recherche affaire
if (Ecran.name='AFPIECECOURS_TIE') or (Ecran.name='AFREGRLIGNE_MUL') then
 begin
 BchangeTiers :=False;
 if ThEdit(GetControl('GL_TIERS'))<>Nil then
   begin
   setControlEnabled ('GL_TIERS',False);
   end;
 if ThEdit(GetControl('GL_AFFAIRE'))<>Nil then
   begin
   setControlEnabled ('GL_AFFAIRE1',False);
   setControlEnabled ('GL_AFFAIRE2',False);
   setControlEnabled ('GL_AFFAIRE3',False);
   setControlEnabled ('GL_AVENANT',False);
   setControlEnabled ('BSELECTAFF1',False);
   setControlEnabled ('BEFFACEAFF1',False);
   end;
 end;
updatecaption(Ecran);
End;

procedure TOF_AFBASEPIECE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('GP_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('GP_TIERS'));
end;

procedure TOF_AFBASEPIECE_MUL.NomsChampsNature(ztable : string; var  App:THEdit
; var TApp:THLabel; var  Viv:TCheckBox);
begin
NatPiece:=Nil;
NatPieceMul:=Nil;
if (ztable = 'GP') then
Begin
  If GetControl('GP_NATUREPIECEG') is ThValComboBox then NatPiece:=THValComboBox(GetControl('GP_NATUREPIECEG'))
    else NatPieceMul:=THMultiValComboBox(GetControl('GP_NATUREPIECEG'));
  App:=THEdit(GetControl('GP_APPORTEUR'));
  TApp:=THLabel(GetControl('TGP_APPORTEUR'));
{$IFNDEF BTP}
  Viv:=TCheckBox(GetControl('GP_VIVANTE'));
{$ENDIF}
End
else
Begin
  If GetControl('GL_NATUREPIECEG') is ThValComboBox then NatPiece:=THValComboBox(GetControl('GL_NATUREPIECEG'))
    else NatPieceMul:=THMultiValComboBox(GetControl('GL_NATUREPIECEG'));
  App:=THEdit(GetControl('GL_APPORTEUR'));
  TApp:=THLabel(GetControl('TGL_APPORTEUR'));
{$IFNDEF BTP}
  Viv:=TCheckBox(GetControl('GL_VIVANTE'));
{$ENDIF}
End
end;


procedure TOF_AFBASEPIECE_MUL.ChargeTobComplete(TobComplete: Tob);
Var Tobpiece : Tob;
    sql :string;
BEGIN
  TobPiece := Tob.Create ('PIECE', TobComplete,-1);
   Sql := '"'+ TFmul(Ecran).Q.FindField('GP_NATUREPIECEG').asstring
   + '";"'+ TFmul(Ecran).Q.FindField('GP_SOUCHE').asstring
   + '";"'+ TFmul(Ecran).Q.FindField('GP_NUMERO').asstring
   + '";"'+ TFmul(Ecran).Q.FindField('GP_INDICEG').asstring + '"';
   TobPiece.selectDB (sql, Nil, False);
END;

procedure TOF_AFBASEPIECE_MUL.AlimPieceVivante;
BEGIN
{$IFNDEF BTP}    //mcd 11/10/02 ajout FRE
if (natpiece.value = 'FPR') or (natpiece.value = 'APR') or (natpiece.value = 'FRE') then
  vivante.checked := true
else
  vivante.checked := false;
{$ENDIF}
END;

procedure TOF_AFBASEPIECE_MUL.OnLoad;
begin
  inherited;
// PL le 11/12/01 : initialisé dans le onload en plus du onargument car si l'utilisateur enregistre un filtre par defaut sur une nature
// puis revient sur le même écran avec une autre nature dans le onargument, le filtre va écraser l'initialisation du onargument
// mcd 10/06/03 si appel depuis client ou mission, il ne faut pas modifier client, même sur recherche affaire
if (NatPiece<>nil) then
if (NatPiece.Enabled = false) or (NatPiece.visible = false) then
    Natpiece.value := sNature;

end;

Initialization
registerclasses([TOF_AFBASEPIECE_MUL]);


end.
