{***********UNITE*************************************************************
Auteur  ...... : CT
Créé le ...... : 06/09/2001
Modifié le ... :   /  /
Description .. : TOF Pour l'écran de lancement d'état des écarts d'inventaire
Mots clefs ... : INVENTAIRE
******************************************************************************}
unit UTOFMBOEdtListeInv;

interface
uses Classes, UTOF, Controls,HCtrls, HEnt1,EntGC,
{$IFDEF EAGLCLIENT}
     UTOB,
{$ELSE}
     DBTables,
{$ENDIF}
     M3FP;

type
    TOF_MBOEdtListeInv = class(TOF)
    private
      FSelectCombo : THValComboBox;
    published
      procedure OnNew; override;
      procedure OnArgument(Arguments : String) ; override ;
      procedure OnUpdate ; override ;   
      procedure OnLoad; override;
    end;

const
    NbRupts = 4;

	// libellés des messages
    TexteMessage: array[1..1] of string 	= (
          {1}        'Le code liste est inexistant'
            );

implementation
uses HMsgBox, buttons, stdctrls, comctrls, Graphics, forms, sysutils;

procedure TOF_MBOEdtListeInv.OnNew;
begin
inherited;
FSelectCombo := THValComboBox(Ecran.FindComponent('KELPRIX'));
FSelectCombo.ItemIndex := 0;
end;

procedure TOF_MBOEdtListeInv.OnArgument(Arguments: String);
var Ctl : TControl ;
    iCol : integer ;
begin
Inherited ;
if (ctxMode in V_PGI.PGIContexte) then
   begin
   SetControlProperty('GIL_DEPOT','Plus','GDE_SURSITE="X"');
   SetControlVisible('GIL_EMPLACEMENT', False);
   SetControlVisible('TGIL_EMPLACEMENT', False);
   end;
// Libellé Dépôt ou Etablissement
if not VH_GC.GCMultiDepots then
   begin
    Ctl := GetControl('TGIL_DEPOT') ;
    if (Ctl <> Nil) and (Ctl is THLabel) then THLabel(Ctl).Caption := 'Etablissement' ;
   end ;
// libellé des familles
   for iCol:=1 to 3 do
    begin
    THLabel(Ecran.FindComponent('TGA_FAMILLENIV'+InttoStr(iCol))).Caption:=RechDom('GCLIBFAMILLE','LF'+InttoStr(iCol),false);
    end;
end;

procedure TOF_MBOEdtListeInv.OnUpdate ;
var QQ : TQuery;
begin
inherited ;
QQ := OpenSQL('SELECT GIE_CODELISTE FROM LISTEINVENT '+
              'WHERE GIE_CODELISTE="'+GetControlText('GIE_CODELISTE')+'"', true);
if QQ.EOF then PgiInfo (TexteMessage[1], Ecran.Caption);
Ferme(QQ) ;
end;

procedure TOF_MBOEdtListeInv.OnLoad;
var QQ : TQuery ;
    i : integer ;
begin
inherited;
while (THCritMaskEdit(Ecran.FindComponent('GIE_CODELISTE')).Text = '') do // Si y'a pas de code liste
begin
  TPageControl(Ecran.FindComponent('Pages')).ActivePage := TTabSheet(Ecran.FindComponent('Standards'));
  with THCritMaskEdit(Ecran.FindComponent('GIE_CODELISTE')) do
   for i := 0 to ComponentCount-1 do
    if Components[i] is TSpeedButton then TSpeedButton(Components[i]).Click; // Click automatique sur l'elipsis
end;                                                                          // = ouverture liste des codes

THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE10')).Text := RechDom('GCINFPRIX', FSelectCombo.Value, true);
QQ:=OpenSql('SELECT CO_LIBRE,CO_ABREGE FROM COMMUN WHERE CO_TYPE="GIP" AND CO_CODE="'+
            FSelectCombo.Value +'"', True);
if not QQ.EOF then
begin
THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE11')).Text:=QQ.Fields[0].AsString ;
THCritMaskEdit(Ecran.FindComponent('XX_VARIABLE12')).Text:=QQ.Fields[1].AsString ;
end;
Ferme(QQ) ;
end;

////////////////////////////////////////////////////////////////////////////////
//****************  CONTROLE DES RUPTURES PARAMETRABLES  ***********************
////////////////////////////////////////////////////////////////////////////////

Procedure TOF_MBOEdtListeInv_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'EDTLISTEINV') and (F.Name <> 'EDTCOMPSTKINV') then exit;
Indice := Integer (Parms[1]);
St_Plus := THEdit(F.FindComponent('CODE')).Text;
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 4 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then
       begin
       St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"' ;
       end;
    END;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := St_Text ;
END;

Procedure TOF_MBOEdtListeInv_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
    Rupt : THValComboBox;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'EDTLISTEINV') and (F.Name <> 'EDTCOMPSTKINV') then exit;
Indice := Integer (Parms[1]);
St_Plus := THEdit(F.FindComponent('CODE')).Text;
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
Rupt:= THValComboBox (F.FindComponent('RUPT'+InttoStr(indice)));
if St_Value = '' then
    BEGIN
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := '';
    Rupt.Plus := St_Plus;
    Rupt.Text := '';
    For i_ind := Indice + 1 to 4 do
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Plus := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := False;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Color := clBtnFace;
        THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
        END;
    END
else
    BEGIN
    if Indice < 4 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Color := clWindow;
        END;
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text :=
            RechDom ('GCGROUPINV', St_Value, True);
    if pos('LF',Rupt.Value) <>0 then
            THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
            RechDom('GCLIBFAMILLE',Rupt.Value,False)
    else    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := Rupt.Text;
    END;
END;

////////////////////////////////////////////////////////////////////////////////

initialization
RegisterClasses([TOF_MBOEdtListeInv]);
RegisterAglProc('EdtListeInv_ChangeGroup', True , 1, TOF_MBOEdtListeInv_ChangeGroup);
RegisterAglProc('EdtListeInv_AffectGroup', True , 1, TOF_MBOEdtListeInv_AffectGroup);
end.
