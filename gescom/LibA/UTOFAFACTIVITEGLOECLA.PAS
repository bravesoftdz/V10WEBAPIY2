{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/07/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFACTIVITEGLOECLA ()
Mots clefs ... : TOF;AFACTIVITEGLOECLA
*****************************************************************}
Unit UTOFAFACTIVITEGLOECLA ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls, ent1,
{$IFDEF EAGLCLIENT}
Maineagl,
{$ELSE}
       db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF, ParamSoc,UTofAfBaseCodeAffaire, HTB97, dicobtp, vierge ;

Type
  TOF_AFACTIVITEGLOECLA = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure bValiderClick(Sender: TObject);
    public
    procedure TypeAppelChargeCle(var AvActionFiche:TActionFiche);override;
  end ;
Function AFLanceFiche_Eclat_Activ(Argument:string):variant;

Implementation

procedure TOF_AFACTIVITEGLOECLA.OnArgument (stArgument : String ) ;
Var
    Critere, Champ, valeur  : String;
    x : integer;
begin
// Recup des critères
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        x:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'TIERS' then SetControlText('ACT_TIERS',Valeur);
        if Champ = 'AFF' then SetControlText('ACT_AFFAIRE',Valeur);
        if Champ = 'DATE' then SetControlText('ACT_DATEACTIVITE',Valeur);
        if Champ = 'ART' then SetControlText('ACT_CODEARTICLE',Valeur);
        if champ = 'RESS' then SetControlText('ACT_RESSOURCE',Valeur);
        if champ = 'STATUT' then SetControlText('STATUTORIG',Valeur);
        if champ = 'MONTANTPV' then THNumEdit(GetControl('MONTANTPVORIG')).Value := strtofloat(Valeur);
        if champ = 'QTE' then THNumEdit(GetControl('QTEORIG')).Value := strtofloat(Valeur);
        if champ = 'VISA' then
            begin
            if (Valeur = 'VIS') then
                TCheckBox(GetControl('CHKVISAORIG')).Checked := true
            else
                TCheckBox(GetControl('CHKVISAORIG')).Checked := false;
            end;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
If Not GetParamSoc ('So_AfAppPoint') then
 begin
 SetControlVisible ('LblVisaLettrage',False);
 SetControlVisible ('LblDevient4',False);
 SetControlVisible ('ChkVisaOrig',False);
 SetControlVisible ('ChkVisaLig1',False);
 SetControlVisible ('ChkVisaLig4',False);
 SetControlVisible ('ChkVisaLig2',False);
 SetControlVisible ('ChkVisaLig3',False);
 end;

// Init du code affaire dans la tof ancêtre
Inherited;

TToolBarButton97(GetControl('BValider')).OnClick := bValiderClick;

ChangeMask(THNumEdit(GetControl('MONTANTPVORIG')), V_PGI.OkDecV, '') ;
ChangeMask(THNumEdit(GetControl('MONTANTPVLIG1')), V_PGI.OkDecV, '') ;
ChangeMask(THNumEdit(GetControl('MONTANTPVLIG2')), V_PGI.OkDecV, '') ;
ChangeMask(THNumEdit(GetControl('MONTANTPVLIG3')), V_PGI.OkDecV, '') ;
ChangeMask(THNumEdit(GetControl('MONTANTPVLIG4')), V_PGI.OkDecV, '') ;
ChangeMask(THNumEdit(GetControl('QTEORIG')), V_PGI.OkDecQ, '') ;
ChangeMask(THNumEdit(GetControl('QTELIG1')), V_PGI.OkDecQ, '') ;
ChangeMask(THNumEdit(GetControl('QTELIG2')), V_PGI.OkDecQ, '') ;
ChangeMask(THNumEdit(GetControl('QTELIG3')), V_PGI.OkDecQ, '') ;
ChangeMask(THNumEdit(GetControl('QTELIG4')), V_PGI.OkDecQ, '') ;
end ;

procedure TOF_AFACTIVITEGLOECLA.bValiderClick(Sender: TObject);
  var
  rep, StatutOrig : string;
  dMontantOrig, dSommeMontant, dQteOrig, dSommeQte : double;
  bPlusieursLignes : boolean;
begin
nextprevcontrol (Ecran);
rep := '';

bPlusieursLignes := (TCheckBox(GetControl('CHKVISALIG2')).State<>cbGrayed)
                    OR (THValComboBox(GetControl('STATUTLIG2')).Value<>'')
                    OR (Valeur(THNumEdit(GetControl('MONTANTPVLIG2')).Text)<>0)
                    OR (Valeur(THNumEdit(GetControl('QTELIG2')).Text)<>0);

// Controle de cohérence des sommes des lignes
StatutOrig := THEdit(GetControl('STATUTORIG')).Text;
dMontantOrig := Valeur(THNumEdit(GetControl('MONTANTPVORIG')).Text);
dQteOrig := Valeur(THNumEdit(GetControl('QTEORIG')).Text);
dSommeMontant := Valeur(THNumEdit(GetControl('MONTANTPVLIG1')).Text) + Valeur(THNumEdit(GetControl('MONTANTPVLIG2')).Text) + Valeur(THNumEdit(GetControl('MONTANTPVLIG3')).Text)  + Valeur(THNumEdit(GetControl('MONTANTPVLIG4')).Text);
dSommeQte := Valeur(THNumEdit(GetControl('QTELIG1')).Text) + Valeur(THNumEdit(GetControl('QTELIG2')).Text) + Valeur(THNumEdit(GetControl('QTELIG3')).Text)  + Valeur(THNumEdit(GetControl('QTELIG4')).Text);

if ((Valeur(THNumEdit(GetControl('MONTANTPVLIG1')).Text)<>0) or bPlusieursLignes) then
if ((dMontantOrig>=0) and ((Valeur(THNumEdit(GetControl('MONTANTPVLIG1')).Text)<0) or (Valeur(THNumEdit(GetControl('MONTANTPVLIG2')).Text)<0) or (Valeur(THNumEdit(GetControl('MONTANTPVLIG3')).Text)<0) or (Valeur(THNumEdit(GetControl('MONTANTPVLIG4')).Text)<0)))
    or (dMontantOrig<0)
 then
    begin
    PGIInfoAf('Attention, l''éclatement de lignes avec des montants négatifs est interdit',Ecran.Caption);
    exit;
    end;
if ((Valeur(THNumEdit(GetControl('QTELIG1')).Text)<>0) or bPlusieursLignes) then
if ((dQteOrig>=0) and ((Valeur(THNumEdit(GetControl('QTELIG1')).Text)<0) or (Valeur(THNumEdit(GetControl('QTELIG2')).Text)<0) or (Valeur(THNumEdit(GetControl('QTELIG3')).Text)<0) or (Valeur(THNumEdit(GetControl('QTELIG4')).Text)<0)))
    or (dQteOrig<0)
 then
    begin
    PGIInfoAf('Attention, l''éclatement de lignes avec des quantités négatives est interdit',Ecran.Caption);
    exit;
    end;

if ((Valeur(THNumEdit(GetControl('MONTANTPVLIG1')).Text)<>0) or bPlusieursLignes) then
if (dSommeMontant<>dMontantOrig) then
    begin
    PGIInfoAf('Attention, la somme des montants PV des lignes éclatées ne correspond pas au montant de la ligne d''origine',Ecran.Caption);
    exit;
    end;

if ((Valeur(THNumEdit(GetControl('QTELIG1')).Text)<>0) or bPlusieursLignes) then
if (dSommeQte<>dQteOrig) then
    begin
    PGIInfoAf('Attention, la somme des quantités des lignes éclatées ne correspond pas à la quantité de la ligne d''origine',Ecran.Caption);
    exit;
    end;

// constitution de la ligne de réponse
if (GetControl('STATUTLIG1').Enabled=true) and ((Trim(GetControlText('STATUTLIG1'))<>'') or (StatutOrig <> '')) then
    begin
      if (THValComboBox(GetControl('STATUTLIG1')).Value<>'') then
          rep := 'STAT1='+THValComboBox(GetControl('STATUTLIG1')).Value+';'
      else
          // PL le 21/10/03 : si on ne choisit pas de statut, on remet celui d'origine par défaut
          rep := 'STAT1='+StatutOrig+';';

      if (THNumEdit(GetControl('MONTANTPVLIG1')).Value<>0) then
          rep := rep + 'MONT1='+floattostr(THNumEdit(GetControl('MONTANTPVLIG1')).Value)+';';
      if (THNumEdit(GetControl('QTELIG1')).Value<>0) then
          rep := rep + 'QTE1='+floattostr(THNumEdit(GetControl('QTELIG1')).Value)+';';
      if (TCheckBox(GetControl('CHKVISALIG1')).state = cbChecked) then
          rep := rep + 'VISA1=VIS;'
      else
      if (TCheckBox(GetControl('CHKVISALIG1')).state = cbunChecked) then
          rep := rep + 'VISA1=ATT;';
    end;

if (GetControl('STATUTLIG2').Enabled=true)
    and ((Trim(GetControlText('STATUTLIG2'))<>'')
          or ((StatutOrig <> '') and ((THNumEdit(GetControl('MONTANTPVLIG2')).Value<>0)
                                        or (THNumEdit(GetControl('QTELIG2')).Value<>0)))) then
    begin
      if (THValComboBox(GetControl('STATUTLIG2')).Value<>'') then
          rep := rep + 'STAT2='+THValComboBox(GetControl('STATUTLIG2')).Value+';'
        else
            // PL le 21/10/03 : si on ne choisit pas de statut, on remet celui d'origine par défaut
          rep := rep + 'STAT2='+StatutOrig+';';

      if (THNumEdit(GetControl('MONTANTPVLIG2')).Value<>0) then
          rep := rep + 'MONT2='+floattostr(THNumEdit(GetControl('MONTANTPVLIG2')).Value)+';';
      if (THNumEdit(GetControl('QTELIG2')).Value<>0) then
          rep := rep + 'QTE2='+floattostr(THNumEdit(GetControl('QTELIG2')).Value)+';';
      if (TCheckBox(GetControl('CHKVISALIG2')).state = cbChecked) then
          rep := rep + 'VISA2=VIS;'
      else
      if (TCheckBox(GetControl('CHKVISALIG2')).state = cbunChecked) then
          rep := rep + 'VISA2=ATT;';
    end;

if (GetControl('STATUTLIG3').Enabled=true)
    and ((Trim(GetControlText('STATUTLIG3'))<>'')
          or ((StatutOrig <> '') and ((THNumEdit(GetControl('MONTANTPVLIG3')).Value<>0)
                                        or (THNumEdit(GetControl('QTELIG3')).Value<>0)))) then
    begin
      if (THValComboBox(GetControl('STATUTLIG3')).Value<>'') then
          rep := rep + 'STAT3='+THValComboBox(GetControl('STATUTLIG3')).Value+';'
        else
            // PL le 21/10/03 : si on ne choisit pas de statut, on remet celui d'origine par défaut
          rep := rep + 'STAT3='+StatutOrig+';';

      if (THNumEdit(GetControl('MONTANTPVLIG3')).Value<>0) then
          rep := rep + 'MONT3='+floattostr(THNumEdit(GetControl('MONTANTPVLIG3')).Value)+';';
      if (THNumEdit(GetControl('QTELIG3')).Value<>0) then
          rep := rep + 'QTE3='+floattostr(THNumEdit(GetControl('QTELIG3')).Value)+';';
      if (TCheckBox(GetControl('CHKVISALIG3')).state = cbChecked) then
          rep := rep + 'VISA3=VIS;'
      else
      if (TCheckBox(GetControl('CHKVISALIG3')).state = cbunChecked) then
          rep := rep + 'VISA3=ATT;';
    end;

if (GetControl('STATUTLIG4').Enabled=true)
    and ((Trim(GetControlText('STATUTLIG4'))<>'')
          or ((StatutOrig <> '') and ((THNumEdit(GetControl('MONTANTPVLIG4')).Value<>0)
                                        or (THNumEdit(GetControl('QTELIG4')).Value<>0)))) then
    begin
      if (THValComboBox(GetControl('STATUTLIG4')).Value<>'') then
          rep := rep + 'STAT4='+THValComboBox(GetControl('STATUTLIG4')).Value+';'
        else
            // PL le 21/10/03 : si on ne choisit pas de statut, on remet celui d'origine par défaut
          rep := rep + 'STAT4='+StatutOrig+';';

      if (THNumEdit(GetControl('MONTANTPVLIG4')).Value<>0) then
          rep := rep + 'MONT4='+floattostr(THNumEdit(GetControl('MONTANTPVLIG4')).Value)+';';
      if (THNumEdit(GetControl('QTELIG4')).Value<>0) then
          rep := rep + 'QTE4='+floattostr(THNumEdit(GetControl('QTELIG4')).Value)+';';
      if (TCheckBox(GetControl('CHKVISALIG4')).state = cbChecked) then
          rep := rep + 'VISA4=VIS;'
      else
      if (TCheckBox(GetControl('CHKVISALIG4')).state = cbunChecked) then
          rep := rep + 'VISA4=ATT;';
    end;

if (rep = '') then
    begin
      PGIInfoAf ('Aucune donnée n''a été modifiée sur la ligne', Ecran.Caption);
      exit;
    end;

TFVierge(Ecran).retour := rep;
Ecran.close;

end;



procedure TOF_AFACTIVITEGLOECLA.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFACTIVITEGLOECLA.TypeAppelChargeCle(var AvActionFiche:TActionFiche);
begin
AvActionFiche := taConsult;
end;

Function AFLanceFiche_Eclat_Activ(argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFACTIVITEGLOECLA','','',argument);
end;

Initialization
  registerclasses ( [ TOF_AFACTIVITEGLOECLA ] ) ;
end.
