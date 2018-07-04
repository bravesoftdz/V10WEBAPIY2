{***********UNITE*************************************************************
Auteur  ...... : CT
Créé le ...... : 06/09/2001
Modifié le ... :   /  /
Description .. : TOF Pour l'écran de lancement d'état des écarts d'inventaire
Mots clefs ... : INVENTAIRE
******************************************************************************}
unit UTOFMBOEdtEcartInv;

interface
uses Classes, UTOF, HCtrls, HEnt1,
{$IFDEF EAGLCLIENT}
     UTOB,
{$ELSE}
     DBTables,
{$ENDIF}
     M3FP;

type
    TOF_MBOEdtEcartInv = class(TOF)
    private

    published
    end;

const
    NbRupts = 4;

	// libellés des messages
    TexteMessage: array[1..1] of string 	= (
          {1}        'Le code liste est inexistant'
            );

implementation
uses HMsgBox, buttons, stdctrls, comctrls, Graphics, forms, sysutils;


////////////////////////////////////////////////////////////////////////////////
//****************  CONTROLE DES RUPTURES PARAMETRABLES  ***********************
////////////////////////////////////////////////////////////////////////////////

Procedure TOF_MBOEdtEcartInv_AffectGroup (parms:array of variant; nb: integer ) ;
var F : TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'EDTECARTINV') then exit;
Indice := Integer (Parms[1]);
St_Plus := THEdit(F.FindComponent('CODE')).Text;
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 3 to 6 do
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

Procedure TOF_MBOEdtEcartInv_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
    Rupt : THValComboBox;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'EDTECARTINV') then exit;
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
    For i_ind := Indice + 1 to 6 do
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
    if Indice < 6 then
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
RegisterClasses([TOF_MBOEdtEcartInv]);
RegisterAglProc('EdtEcartInv_ChangeGroup', True , 1, TOF_MBOEdtEcartInv_ChangeGroup);
RegisterAglProc('EdtEcartInv_AffectGroup', True , 1, TOF_MBOEdtEcartInv_AffectGroup);
end.
