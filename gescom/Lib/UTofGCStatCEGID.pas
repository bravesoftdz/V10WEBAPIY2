unit UTofGCStatCEGID;

interface
uses  M3FP, StdCtrls,Controls,Classes, Graphics, forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,uTofAfBaseCodeAffaire,

{$IFDEF EAGLCLIENT}
      MaineAGL,eQRS1,
{$ELSE}
      Fe_Main,db,dbTables,MajTable,QRS1,
{$ENDIF}
      UTOF, AglInit, Agent, Ent1,EntGC ;

Type
     TOF_StatCEGID = Class (TOF_AFBASECODEAFFAIRE)
       private
       stNatureEtat : string ;

       public
       procedure OnArgument (stArgument : String )  ; override ;
       procedure OnUpdate ; override ;
       procedure OnNew ; override;
       procedure OnLoad ; override ;
       procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
     END ;

Const
	// libellés des messages
	TexteMessage: array[1..1] of string 	= (
          {1}         'Vous devez sélectionner une nature de pièce.'
                 );
implementation

/////////////////////////////////////////////////////////////////////////////
procedure TOF_StatCEGID.OnArgument (stArgument : String ) ;
Var stArg,stVenteAchat, St_Plus : string ;
    CC : THValComboBox ;
    i : integer;
begin
    Inherited;
stArg := stArgument;
stVenteAchat := ReadTokenSt(stArg); stNatureEtat := ReadTokenSt(stArg);
St_Plus:='GPP_VENTEACHAT="'+stVenteAchat+'"';
if not(ctxScot in V_PGI.PGIContexte) then THValComboBox (Ecran.FindComponent('NATUREPIECEG')).Plus := St_Plus ;
THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(0,'Facture clients');
THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(0,'FAC');
THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(1,'Avoir clients');
THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(1,'ZZ2');
THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(2,'Facture + Avoir clients');
THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(2,'ZZ1');

if VH_GC.GCIfDefCEGID then
begin
  SetControlText('REFMARGE','DPR');
  SetControlEnabled('REFMARGE',false);
end
else
begin
  SetControlCaption('RPRIXACHAT','Prix d''achat');
  SetControlText('REFMARGE',VH_GC.GCMargeArticle);
end;

SetControlText('NATUREPIECEG','ZZ1');
THValComboBox (Ecran.FindComponent('GL_DEPOT')).ItemIndex := 0;
THValComboBox (Ecran.FindComponent('GP_ETABLISSEMENT')).ItemIndex := 0;
if stNatureEtat = '' then
   begin
   THValComboBox (Ecran.FindComponent('RUPT1')).ItemIndex := 0;
   //critére de regroupement affaire caché si pas de contexte affaire PA
   if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
      begin
      for i := 1 to 6 do begin SetControlProperty('RUPT'+IntToStr(i),'Plus','AND CO_CODE<>"AFF"'); end;
      end;
   end;

if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
      begin SetControlVisible ('TGP_AFFAIRE', False); SetControlVisible ('TGP_AFFAIRE_', False); end;

CC:=THValComboBox(GetControl('GP_DOMAINE')) ; if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
end;

Procedure TOF_StatCEGID.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('GP_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ; Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ; Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
Aff_:=THEdit(GetControl('GP_AFFAIRE_'))   ; Aff0_:=Nil ;
Aff1_:=THEdit(GetControl('GP_AFFAIRE1_')) ; Aff2_:=THEdit(GetControl('GP_AFFAIRE2_')) ;
Aff3_:=THEdit(GetControl('GP_AFFAIRE3_')) ; Aff4_:=THEdit(GetControl('GP_AVENANT_'))  ;
END ;

procedure TOF_StatCEGID.OnNew ;
begin
    Inherited;
if stNatureEtat='' then exit;
SetControlText('FEtat',stNatureEtat);
Ecran.Caption := Ecran.Caption + ' : ' + THValComboBox(GetControl('FEtat')).Text;
UpdateCaption(Ecran) ;
SetControlEnabled('FEtat',False);
TFQRS1(Ecran).FCodeEtat:=GetControlText('FEtat');
end;

procedure TOF_StatCEGID.OnLoad;
var stWhere : string;
begin
if stNatureEtat <> '' then SetControlText('FEtat',stNatureEtat);
    Inherited;
if GetControlText('NATUREPIECEG') = '' then
    BEGIN
    SetControlText('XX_WHERE', 'GP_NATUREPIECEG=""');
    LastError:=1;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    END;
stWhere := '';
SetControlText('GL_ARTICLE',GetControlText ('CODEARTICLE'));
if GetControlText ('CODEARTICLE_') <> '' then
     SetControlText('GL_ARTICLE_',format('%-18.18s%16.16s',[GetControlText ('CODEARTICLE_'),StringOfChar('Z',16)]))
else SetControlText('GL_ARTICLE_','');

if GetControlText ('NATUREPIECEG') <> '' then
    begin
    if GetControlText('NATUREPIECEG')='ZZ1' then
       stWhere := ' AND (GP_NATUREPIECEG="FAC" OR GP_NATUREPIECEG="AVC" OR GP_NATUREPIECEG="AVS") '
       else if GetControlText('NATUREPIECEG')='ZZ2' then
            stWhere := ' AND (GP_NATUREPIECEG="AVC" OR GP_NATUREPIECEG="AVS") '
            else stWhere := ' AND GP_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
    end;
SetControlText('XX_VARIABLE7',GetControlText('REFMARGE'));
SetControlText('XX_WHERE',stWhere);
SetControlText('XX_ORDERBY', 'GP_SOUCHE,GP_NUMERO');

if TRadioButton(GetControl ('RPRIXACHAT')).checked = True then
     SetControlText ('XX_VARIABLE11', GetControlText('REFMARGE'))
else SetControlText ('XX_VARIABLE11', 'Qté');

end;

procedure TOF_StatCEGID.OnUpdate;
begin
    Inherited;

end;

/////////////////////////////////////////////////////////////////////////////
Procedure GCStatCEGID_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St, PlusS3 : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCSTATCEGID') then exit;
Indice := Integer (Parms[1]);
St_Plus := '';
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 6 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"';
    END;
{$IFDEF CCS3}
PlusS3:='AND CO_CODE<>"AFF" AND (CO_CODE<"LT5" OR CO_CODE>"LTA")' ;
St_Plus:=St_Plus+PlusS3 ;
{$ENDIF}
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
if St_Value<>'' then THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value
                else THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).ItemIndex := 0;
END;

Procedure GCStatCEGID_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
    Entete, Detail : boolean ;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCSTATCEGID') then exit;
Indice := Integer (Parms[1]);
St_Plus := '';
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
if St_Value = '' then
    BEGIN
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := '';
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked := False;
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := False;
    For i_ind := Indice + 1 to 6 do
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := False;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Color := clBtnFace;
        THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := False;
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := False;
        END;
    END else
    BEGIN
    if Indice < 6 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Color := clWindow;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).ItemIndex := 0;
        END;
    Entete := TCheckBox (F.FindComponent('ENTETEPIECE')).Checked;
    Detail := TCheckBox (F.FindComponent('DETAILLIGNE')).Checked;
    if (Entete) OR (Detail) then
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := True;
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text :=
            RechDom ('GCGROUPSTATCEG', St_Value, True);
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
            string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
    END;
END;

Procedure GCStatCEGID_AffectSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Value : string;
    i_ind : integer;
    Entete, Detail : boolean ;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCSTATCEGID') then exit;
Entete := TCheckBox (F.FindComponent('ENTETEPIECE')).Checked;
Detail := TCheckBox (F.FindComponent('DETAILLIGNE')).Checked;
if (not Entete) AND (not Detail) then
    BEGIN
    For i_ind := 1 to 6 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := False;
    END else
    BEGIN
    For i_ind := 1 to 6 do
        BEGIN
        St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
        if St_Value <> '' then
            TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := True;
        END;
    END;
END;

Procedure GCStatCEGID_ChangeSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    i_ind, Indice : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCSTATCEGID') then exit;
Indice := Integer (Parms[1]);
if TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked then
    BEGIN
    For i_ind := 1 to Indice - 1 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := True;
    END else
    BEGIN
    For i_ind := Indice + 1 to 6 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := False;
    END;
END;

/////////////////////////////////////////////////////////////////////////////

procedure InitTOFGCStatCEGID ();
begin
RegisterAglProc('GCStatCEGID_ChangeGroup', True , 1, GCStatCEGID_ChangeGroup);
RegisterAglProc('GCStatCEGID_AffectGroup', True , 1, GCStatCEGID_AffectGroup);
RegisterAglProc('GCStatCEGID_AffectSautPage', True , 0, GCStatCEGID_AffectSautPage);
RegisterAglProc('GCStatCEGID_ChangeSautPage', True , 0, GCStatCEGID_ChangeSautPage);
end;

Initialization
registerclasses([TOF_StatCEGID]) ;
InitTOFGCStatCEGID ();
end.

 