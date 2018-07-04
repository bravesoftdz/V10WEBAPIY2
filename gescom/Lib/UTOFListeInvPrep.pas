{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 5/05/2000
Modifié le ... :   /  /
Description .. : TOF Pour l'écran de lancement d'état de liste
Suite ........ : préparatoire d'inventaire
Mots clefs ... : INVENTAIRE;PREPARATOIRE;STOCK
*****************************************************************}
unit UTOFListeInvPrep;

interface
uses Classes, UTOF, HCtrls, HEnt1,
{$IFDEF EAGLCLIENT}
     UTOB,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     M3FP;

type
    TOF_ListeInvPrep = class(TOF)
    private

    published
      procedure OnNew; override;
      procedure OnArgument(Arguments : String) ; override ;
      procedure OnUpdate ; override ;
      procedure OnLoad; override;
      procedure OnComboChange(Sender : TObject);
      procedure OnComboPopUp(Sender : TObject);
    end;

const
    NbRupts = 4;

	// libellés des messages
    TexteMessage: array[1..1] of string 	= (
          {1}        'Le code liste est inexistant'
            );

implementation
uses HMsgBox, buttons, stdctrls, comctrls, Graphics, forms, sysutils;

procedure TOF_ListeInvPrep.OnNew;
var i : integer;
begin
inherited;
for i := 1 to NbRupts do
  begin
  THValComboBox(GetControl('RUPT'+inttostr(i))).OnChange := OnComboChange;
  THValComboBox(GetControl('RUPT'+inttostr(i))).OnDropDown := OnComboPopup;
  THLabel(GetControl('TGA_FAMILLENIV'+inttostr(i))).Caption := RechDom('GCLIBFAMILLE', 'LF'+inttostr(i), false)
  end;
end;

procedure TOF_ListeInvPrep.OnArgument(Arguments: String);
begin
Inherited ;
if (ctxMode in V_PGI.PGIContexte) then
   begin
   SetControlVisible('GIL_EMPLACEMENT', False);
   SetControlVisible('TGIL_EMPLACEMENT', False);
   SetControlVisible('CHKLOTS', False);
   SetControlChecked('CHKLOTS', False);
   end;
end;

procedure TOF_ListeInvPrep.OnUpdate ;
var QQ : TQuery;
begin
inherited ;
QQ := OpenSQL('SELECT GIE_CODELISTE FROM LISTEINVENT '+
              'WHERE GIE_CODELISTE="'+GetControlText('GIE_CODELISTE')+'"', true);
if QQ.EOF then PgiInfo (TexteMessage[1], Ecran.Caption);
Ferme(QQ) ;
end;

procedure TOF_ListeInvPrep.OnLoad;
var i : integer;
begin
inherited;
if GetControlText('GIE_CODELISTE') = '' then // Si y'a pas de code liste
begin
  TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl('Standards'));
  with THCritMaskEdit(GetControl('GIE_CODELISTE')) do
   for i := 0 to ComponentCount-1 do
    if Components[i] is TSpeedButton then TSpeedButton(Components[i]).Click; // Click automatique sur l'elipsis
end;                                                                          // = ouverture liste des codes

if GetControlText('GIE_CODELISTE') = ''   // Si on a fait cancel
  then SetControlText('XX_WHERE','GIE_CODELISTE=""')  // Pas d'etat
  else SetControlText('XX_WHERE',''); // sinon ok
end;


procedure TOF_ListeInvPrep.OnComboChange(Sender : TObject);
var RuptNb : String;
    FCombo : THValComboBox;
    NextComboName : String;
begin
FCombo := (Sender as THValComboBox);
RuptNb := Copy(FCombo.Name,5,1);

if FCombo.Value = '' then
  begin
  SetControlText('XX_RUPTURE'+RuptNb,'');
  SetControlText('XX_VARIABLE'+RuptNb,'');
  end else
  begin
  SetControlText('XX_RUPTURE'+RuptNb,RechDom(FCombo.DataType, FCombo.Value, true));
  if (FCombo.Value = 'LF1') or (FCombo.Value = 'LF2') or (FCombo.Value = 'LF3')
    then SetControlText('XX_VARIABLE'+RuptNb,RechDom('GCLIBFAMILLE', FCombo.Value, false))
    else SetControlText('XX_VARIABLE'+RuptNb,RechDom(FCombo.DataType, FCombo.Value, false));
  end;

NextComboName := FCombo.Name; NextComboName[5] := Chr(Ord(NextComboName[5])+1);
if (GetControl(NextComboName) <> nil) and (FCombo.Value <> '') then
  begin
  SetControlEnabled(NextComboName,true);
  SetControlProperty(NextComboName,'Color',clWindow);
  end else
  while GetControl(NextComboName) <> nil do
    begin
    with THValComboBox(GetControl(NextComboName)) do
      begin Enabled := False; Color := clBtnFace; Value := ''; end;
    NextComboName[5] := Chr(Ord(NextComboName[5])+1);
    end;
end;

// Y'a un bug quand on passe par le OnEnter via le script : si on remet à <<Aucun>>
// une rupture alors que plusieurs suivantes sont sélectionnées, et qu'on décide de
// remplacer ensuite le <<Aucun>> par autre chose, le contenu du combo est pas mis
// à jour puisque le OnEnter n'est pas rééxécuté
procedure TOF_ListeInvPrep.OnComboPopUp(Sender : TObject);
var St_Plus, St_Value, St_Text, St : String;
    FCombo : THValComboBox;
    Indice, i_ind : Integer;
begin
FCombo := (Sender as THValComboBox);
Indice := StrToInt(Copy(FCombo.Name,5,1));
St_Plus := '';
St_Value := FCombo.Value;
St_Text := FCombo.Text;
for i_ind := 1 to NbRupts do
  begin
  if i_ind = Indice then Continue;
  St := String(THValComboBox(GetControl('RUPT'+IntToStr(i_ind))).Value);
  if St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"'+St+'"';
  end;
FCombo.Plus := St_Plus;
FCombo.Value := St_Value;
FCombo.Text := St_Text;
end;


initialization
RegisterClasses([TOF_ListeInvPrep]);

end.
