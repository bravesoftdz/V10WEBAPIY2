unit UtofSuiviActivite;

interface
uses  StdCtrls,Classes, UtilGC,
{$IFDEF EAGLCLIENT}
     UTOB,
{$ELSE}
     db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
      forms,sysutils,
      HCtrls,HEnt1,UTOF,
      M3FP,Graphics,UtilSelection,UtilRT,ParamSoc,utofAfBaseCodeAffaire,EntGC;

Type
    TOF_SuiviActivite = Class (TOF_AFBASECODEAFFAIRE)
     private
         stProduitpgi : string;
{$IFDEF AFFAIRE}
     protected
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override;
        procedure bSelectAff1Click(Sender: TObject);     override ;
{$ENDIF AFFAIRE}
     public
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad ; override ;

     END;

implementation
{$ifdef GIGI}
Var     Plus : string;
{$endif}

procedure TOF_SuiviActivite.OnArgument(stArgument : String );
var F : TForm;
    i : integer;
begin
inherited ;
  stProduitpgi := stArgument;
  if stProduitpgi = '' then stProduitpgi := 'GRC';
  if stArgument <> 'GRF' then
    begin
    F := TForm (Ecran);
    MulCreerPagesCL(F,'NOMFIC=GCTIERS');

    if GetParamSocSecur('SO_RTGESTINFOS001',False) = True then
        MulCreerPagesCL(F,'NOMFIC=RTACTIONS');
    for i:=1 to 3 do
        SetControlCaption('TRAC_TABLELIBRE'+IntToStr(i),'&'+RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE)) ;
    end
    else
    for i:=1 to 3 do
        SetControlCaption('TRAC_TABLELIBREF'+IntToStr(i),'&'+RechDom('RTLIBCHAMPSLIBRES','FA'+IntToStr(i),FALSE)) ;
{$IFDEF AFFAIRE}
if ( not (ctxAffaire in V_PGI.PGIContexte) ) and ( not ( ctxGCAFF in V_PGI.PGIContexte) ) then
{$ENDIF}
      begin
      SetControlVisible ('BEFFACEAFF1',false); SetControlVisible ('BSELECTAFF1',false);
      SetControlVisible ('TRAC_AFFAIRE',false); SetControlVisible ('RAC_AFFAIRE1',false);
      SetControlVisible ('RAC_AFFAIRE2',false); SetControlVisible ('RAC_AFFAIRE3',false);
      SetControlVisible ('RAC_AVENANT',false);
      end;
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then  SetControlVisible ('PRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    end;
  end;
{$Ifdef GIGI}
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 plus := ' AND CO_CODE <>"COM" AND CO_CODE<>"ZON"';
 if (GetControl('RAC_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    plus := plus+' AND CO_CODE<>"OPE"';
    end;
 if (GetControl('RAC_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False)
    then  begin
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    plus := plus+' AND CO_CODE<>"PRO"';
    end;
 If (Not VH_GC.GaSeria) or not (GetParamSocSecur ('SO_AFRTPROPOS',False)) then
   begin
   plus := plus+' AND CO_CODE<>"AFF"';
   if (GetControl('TRAC_AFFAIRE') <> nil) then  SetControlVisible('TRAC_AFFAIRE',false);
   if (GetControl('RAC_AFFAIRE1') <> nil) then  SetControlVisible('RAC_AFFAIRE1',false);
   if (GetControl('RAC_AFFAIRE2') <> nil) then  SetControlVisible('RAC_AFFAIRE2',false);
   if (GetControl('RAC_AFFAIRE3') <> nil) then  SetControlVisible('RAC_AFFAIRE3',false);
   if (GetControl('RAC_AVENANT') <> nil) then  SetControlVisible('RAC_AVENANT',false);
   if (GetControl('BEFFACEAFF1') <> nil) then  SetControlVisible('BEFFACEAFF1',false);
   if (GetControl('BSELECTAFF1') <> nil) then  SetControlVisible('BSELECTAFF1',false);
   end;
 If (Not GetParamSocSecur ('SO_AFRTCHAINAGE',False)) then
   begin
   SetControlVisible('RAC_NUMCHAINAGE',false);
   SetControlVisible('TRAC_NUMCHAINAGE',false);
   end;
{$endif}
{$IFDEF GRCLIGHT}
  if ( stProduitpgi = 'GRC') and (not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False)) then
    begin
    SetControlVisible('OPERATIONACT',false);
    SetControlVisible('TOPERATIONACT',false);
    end;
{$ENDIF GRCLIGHT}

end;

procedure TOF_SuiviActivite.OnLoad;
Var AdrRed   : TRadioButton;
    AdrComp  : TRadioButton;
    Confid : string;
begin
Inherited;
if GetControlText ('XX_RUPTURE1') <> '' then SetControlText('XX_ORDERBY',  'RAC_AUXILIAIRE')
else SetControlText('XX_ORDERBY',  'RAC_DATEACTION,RAC_AUXILIAIRE');
AdrRed := TRadioButton(GetControl('CADRRED'));
AdrComp := TRadioButton(GetControl('CADRCOMP'));
if (AdrRed.Checked=True) or (AdrComp.Checked=True) then
   begin
   SetControlText ('TITREADR',TraduireMemoire('Adresse'));
   end else
   begin
   SetControlText ('TITREADR',TraduireMemoire('Raison sociale'));
   end;

  if stProduitpgi = 'GRC' then Confid:='CON' else Confid:='CONF';
  SetControlText('XX_WHERE',RTXXWhereConfident(Confid)) ;
end;

{$IFDEF AFFAIRE}
procedure TOF_SuiviActivite.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('RAC_AFFAIRE'));
Aff1:=THEdit(GetControl('RAC_AFFAIRE1'));
Aff2:=THEdit(GetControl('RAC_AFFAIRE2'));
Aff3:=THEdit(GetControl('RAC_AFFAIRE3'));
Aff4:=THEdit(GetControl('RAC_AVENANT'));
Tiers:=THEdit(GetControl('RAC_TIERS'));
end;

procedure TOF_SuiviActivite.bSelectAff1Click(Sender: TObject);
begin
    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4);
end;
{$ENDIF AFFAIRE}

Procedure TOF_SuiviActivite_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'RTACTIONS_ETAT2') and (F.Name <> 'RFACTIONS_ETAT') then exit;
Indice := Integer (Parms[1]);
St_Plus := '';
if not (ctxaffaire in V_PGI.PGIContexte) then St_Plus := 'AND CO_CODE not Like "R%"';
{$ifdef GIGI}
St_Plus := plus;
{$endif}
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 6 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"';
    END;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
if St_Value <> '' then
    begin
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value;
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := St_Text ;    //cd 200401
    end else
    begin
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).ItemIndex := 0;
    end;
END;

Procedure TOF_SuiviActivite_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
    Q : TQuery;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'RTACTIONS_ETAT2') and (F.Name <> 'RFACTIONS_ETAT') then exit;
Indice := Integer (Parms[1]);
St_Plus := '';
{$ifdef GIGI}
St_Plus := plus;
{$endif}
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
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Text := '';      //cd 200401
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
        END;
{    Detail := TCheckBox (F.FindComponent('DETAILLIGNE')).Checked;
    if (Detail) then
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := True;  }
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := True;  
    if (F.Name = 'RTACTIONS_ETAT2') then
       Q := OpenSQL ('SELECT CO_LIBRE FROM COMMUN '+
         'WHERE CO_TYPE="RGA" AND CO_CODE="'+ St_Value+'"', true)
    else
       Q := OpenSQL ('SELECT CO_LIBRE FROM COMMUN '+
         'WHERE CO_TYPE="RGF" AND CO_CODE="'+ St_Value+'"', true);

    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text :=
//            RechDom ('RTGROUPACTIONS', St_Value, True);
              Q.Fields[0].AsString;
    Ferme (Q);
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
            string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
    END;
END;

/////////////////////////////////////////////////////////////////////////////
Procedure TOF_SuiviActivite_AffectSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Value : string;
    i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'RTACTIONS_ETAT2') and (F.Name <> 'RFACTIONS_ETAT') then exit;
{Detail := TCheckBox (F.FindComponent('DETAILLIGNE')).Checked;
if (not Detail) then
    BEGIN
    For i_ind := 1 to 6 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := False;
    END else
    BEGIN  }
    For i_ind := 1 to 6 do
        BEGIN
        St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
        if St_Value <> '' then
            TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := True;
        END;
   { END;}
END;

/////////////////////////////////////////////////////////////////////////////
Procedure TOF_SuiviActivite_ChangeSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    i_ind, Indice : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'RTACTIONS_ETAT2') and (F.Name <> 'RFACTIONS_ETAT') then exit;
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

Initialization
registerclasses([TOF_SuiviActivite]);
RegisterAGLProc('SuiviActivite_ChangeGroup', True, 1, TOF_SuiviActivite_ChangeGroup);
RegisterAGLProc('SuiviActivite_AffectGroup', True, 1, TOF_SuiviActivite_AffectGroup);
RegisterAGLProc('SuiviActivite_ChangeSautPage', True, 1, TOF_SuiviActivite_ChangeSautPage);
RegisterAGLProc('SuiviActivite_AffectSautPage', True, 1, TOF_SuiviActivite_AffectSautPage);
end. 
