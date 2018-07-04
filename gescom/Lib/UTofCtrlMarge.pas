unit UTofCtrlMarge;

interface                                                                                  

uses  M3FP, StdCtrls,Controls,Classes, Graphics, forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      dbTables,Fe_Main, db,MajTable,
{$ENDIF}
      UTOF, AglInit, Agent, Entgc;

Type
     TOF_CtrlMarge = Class (TOF)
       private

       public
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;
       procedure OnNew ; override ;
       procedure OnArgument(stArgument : String ) ; override ;       
     END ;

Const
	// libellés des messages
	TexteMessage: array[1..2] of string 	= (
          {1}         'Vous devez sélectionner une nature de pièce.'
          {2}        ,'La plage de sélection de dates ne peut excéder 31 jours'
                 );


implementation

// Initialisation Combos de sélection
procedure TOF_CtrlMarge.OnArgument (stArgument : String ) ;
var Iind : integer;
begin
     Inherited;
     THValComboBox (GetControl('GP_ETABLISSEMENT')).ItemIndex := 0;
     THValComboBox (GetControl('GP_DEPOT')).ItemIndex := 0;
     THValComboBox (GetControl('RUPT1')).ItemIndex := 0;
     Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FAC');
     if Iind<0 then Iind:=0;
     THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Avoir clients');
     THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ1');
     SetControlText('NATUREPIECEG','ZZ1');
     for iInd:=1 to 6 do
   //THValComboBox(GetControl('REGROUP'+intToStr(iInd))).Plus:='';
   if (not VH_GC.GCMultiDepots) then
        THValComboBox (TForm(Ecran).FindComponent ('RUPT'+InttoStr (iInd))).Plus:=' AND CO_CODE<>"DEP"';
end;

// Positionnement Var "Référence marge" et "Libellés Familles"
procedure TOF_CtrlMarge.OnNew;
var st_Etiq : string;
begin
    Inherited;
    SetControlText('REFMARGE', VH_GC.GCMARGEARTICLE);
    St_Etiq := RechDom ('GCLIBFAMILLE', 'LF1', false);
    if St_Etiq<>'' then
       begin
       SetControlText('TGA_FAMILLENIV1', St_Etiq);
       end else
       begin
       SetControlText('TGA_FAMILLENIV1', '');
       SetControlEnabled('GA_FAMILLENIV1', false);
       SetControlVisible('TGA_FAMILLENIV1', false);
       SetControlVisible('GA_FAMILLENIV1', false);
       end;
    St_Etiq := RechDom ('GCLIBFAMILLE', 'LF2', false);
    if St_Etiq<>'' then
       begin
       SetControlText('TGA_FAMILLENIV2', St_Etiq);
       end else
       begin
       SetControlText('TGA_FAMILLENIV2', '');
       SetControlEnabled('GA_FAMILLENIV2', false);
       SetControlVisible('TGA_FAMILLENIV2', false);
       SetControlVisible('GA_FAMILLENIV2', false);
       end;
    St_Etiq := RechDom ('GCLIBFAMILLE', 'LF3', false);
    if St_Etiq<>'' then
       begin
       SetControlText('TGA_FAMILLENIV3', St_Etiq);
       end else
       begin
       SetControlText('TGA_FAMILLENIV3', '');
       SetControlEnabled('GA_FAMILLENIV3', false);
       SetControlVisible('TGA_FAMILLENIV3', false);
       SetControlVisible('GA_FAMILLENIV3', false);
       end;

// Maj de l'écran en fonction du paramétrage mono-dépôt ou multi-dépôts
if not VH_GC.GCMultiDepots then
   begin
   SetControlVisible('TGP_DEPOT', false);
   SetControlVisible('GP_DEPOT', false);
   end;
end;
/////////////////////////////////////////////////////////////////////////////
procedure TOF_CtrlMarge.OnLoad;
Var stWhere : string ;
begin
    Inherited;
stWhere := '';
if GetControlText('NATUREPIECEG') = '' then
    BEGIN
    SetControlText('XX_WHERE_NAT', 'GP_NATUREPIECEG=""');
    LastError:=1 ; LastErrorMsg:=TexteMessage[LastError];
    AfficheError:=False;
    PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
    exit;
    END ELSE
    BEGIN
    if GetControlText('NATUREPIECEG')='ZZ1' then
         stWhere := ' AND (GP_NATUREPIECEG="FAC" OR GP_NATUREPIECEG="AVC" OR GP_NATUREPIECEG="AVS") '
    else stWhere := ' AND GP_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
    END;
SetControlText('XX_WHERE_NAT',stWhere);
if VH_GC.GCIfDefCEGID then
  if (StrToDate(GetControlText('GP_DATEPIECE_')) - StrToDate(GetControlText('GP_DATEPIECE'))) > 31  then
    BEGIN
    SetActiveTabsheet('STANDARDS') ; SetFocusControl ('GP_DATEPIECE');
    SetControlText('GP_DATEPIECE',  DateToStr(V_PGI.DateEntree));
    SetControlText('GP_DATEPIECE_', DateToStr(V_PGI.DateEntree));
    LastError:=2 ; LastErrorMsg:=TexteMessage[LastError];
    AfficheError:=False;
    PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
    exit;
    END;
SetControlText('XX_ORDERBY', 'GP_SOUCHE,GP_NUMERO');
end;
///////////////////////////////////////////////////////////////////////////////
procedure TOF_CtrlMarge.OnUpdate;
begin
    Inherited;
if GetControlText('NATUREPIECEG') = '' then
    begin
    SetControlText('XX_WHERE_NAT', '');
    end;
if GetControlText('REFMARGE') = 'DPA' then
    begin
    SetControlText('XX_VARIABLE10', 'GL_DPA');
    end;
if GetControlText('REFMARGE') = 'DPR' then
    begin
    SetControlText('XX_VARIABLE10', 'GL_DPR');
    end;
if GetControlText('REFMARGE') = 'PMA' then
    begin
    SetControlText('XX_VARIABLE10', 'GL_PMAP');
    end;
if GetControlText('REFMARGE') = 'PMR' then
    begin
    SetControlText('XX_VARIABLE10', 'GL_PMRP');
    end;
end;

/////////////////////////////////////////////////////////////////////////////
Procedure TOF_CtrlMarge_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCCTRLMARGE') then exit;
Indice := Integer (Parms[1]);
if (not VH_GC.GCMultiDepots) then St_Plus := ' AND CO_CODE<>"DEP"'
else St_Plus :='';//THValComboBox (F.FindComponent('RUPT1')).Plus;
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 4 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"';
    END;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
if St_Value <> '' then
    begin
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value;
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := St_Text;
    end else
    begin
    THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).ItemIndex := 0;
    end;
END;
//////////////////////////////////////////////////////////////////////////////
Procedure TOF_CtrlMarge_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
    Entete, Detail : boolean;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCCTRLMARGE') then exit;
Indice := Integer (Parms[1]);
if (not VH_GC.GCMultiDepots) then St_Plus := ' AND CO_CODE<>"DEP"'
else St_Plus :='';//THValComboBox (F.FindComponent('RUPT1')).Plus;
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
if St_Value = '' then
    BEGIN
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := '';
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked := False;
    TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := False;
    THValComboBox (F.FindComponent('RUPT1')).Plus := St_Plus;
    For i_ind := Indice + 1 to 4 do
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Text := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Color := clBtnFace;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := False;
        THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := False;
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := False;
        END;
    END else
    BEGIN
    if Indice < 4 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Color := clWindow;
        //THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).ItemIndex := 0;
        END;
    Entete := TCheckBox (F.FindComponent('ENTPIECE')).Checked;
    Detail := TCheckBox (F.FindComponent('DETLIGNE')).Checked;
    if (Entete) OR (Detail) then
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Enabled := True;
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text :=
            RechDom ('GCGROUPCTRLMARGE', St_Value, True);
    if (St_value = 'LF1') or (St_value = 'LF2') or (St_value = 'LF3') then
      begin
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
          RechDom ('GCLIBFAMILLE', St_value, false);
      end else
      begin
         THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
           string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
      end;
    END;
END;
/////////////////////////////////////////////////////////////////////////////
Procedure TOF_CtrlMarge_AffectSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Value : string;
    i_ind : integer;
    Entete, Detail : boolean ;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCCTRLMARGE') then exit;
Entete := TCheckBox (F.FindComponent('ENTPIECE')).Checked;
Detail := TCheckBox (F.FindComponent('DETLIGNE')).Checked;
if (not Entete) AND (not Detail) then
    BEGIN
    For i_ind := 1 to 4 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := False;
    END else
    BEGIN
    For i_ind := 1 to 4 do
        BEGIN
        St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
        if St_Value <> '' then
            TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Enabled := True;
        END;
    END;
END;
/////////////////////////////////////////////////////////////////////////////
Procedure TOF_CtrlMarge_ChangeSautPage (parms:array of variant; nb: integer ) ;
var F: TForm;
    i_ind, Indice : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCCTRLMARGE') then exit;
Indice := Integer (Parms[1]);
if TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(Indice))).Checked then
    BEGIN
    For i_ind := 1 to Indice - 1 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := True;
    END else
    BEGIN
    For i_ind := Indice + 1 to 4 do
        TCheckBox (F.FindComponent('SAUTRUPT'+InttoStr(i_ind))).Checked := False;
    END;
END;

/////////////////////////////////////////////////////////////////////////////
procedure InitTOFCtrlMarge();
begin
RegisterAGLProc('CtrlMarge_ChangeGroup', True, 1, TOF_CtrlMarge_ChangeGroup);
RegisterAGLProc('CtrlMarge_AffectGroup', True, 1, TOF_CtrlMarge_AffectGroup);
RegisterAGLProc('CtrlMarge_ChangeSautPage', True, 1, TOF_CtrlMarge_ChangeSautPage);
RegisterAGLProc('CtrlMarge_AffectSautPage', True, 1, TOF_CtrlMarge_AffectSautPage);
end;

Initialization
registerclasses([TOF_CtrlMarge]) ;
InitTOFCtrlMarge();

end.
