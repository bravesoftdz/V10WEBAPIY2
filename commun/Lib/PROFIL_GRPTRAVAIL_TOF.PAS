{***********UNITE*************************************************
Auteur  ...... : G.Harlez
Créé le ...... : 16/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PROFIL_GRPTRAVAIL ()
Mots clefs ... : TOF;PROFIL_GRPTRAVAIL
*****************************************************************}
Unit PROFIL_GRPTRAVAIL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     VIERGE;

Type
  TOF_PROFIL_GRPTRAVAIL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Obj_Combo      : THValComboBox;
    Obj_ChkbAllPrf : THCheckbox;
    Obj_ChkbNoPrf  : THCheckbox;
    Obj_BtnErase   : TToolbarButton97;

    s_mode : string;

    procedure DoClick  (Sender : Tobject);
    procedure DoChange (Sender : Tobject);
  end ;

Implementation

uses TntWideStrings;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PROFIL_GRPTRAVAIL.OnUpdate ;
var
  s_profil,s_allprf,s_noprf : string;
begin
  Inherited ;
  s_profil := '';
  s_allprf := '-';
  s_noprf  := '-';

  if Obj_Combo.ItemIndex <> -1 then
    s_profil := Obj_Combo.Values.Strings[Obj_Combo.ItemIndex];

  if Obj_ChkbAllPrf.Checked then
    s_allprf := 'X';

  if Obj_ChkbNoPrf.Checked then
    s_noprf := 'X';

  if s_mode <> 'MODIFICATION' then
    // Valeur de retour sous la forme : profil|affectation à tous (X ou -)|aucune affectation (X ou -)
    TFVierge(ecran).Retour := s_profil+'|'+s_allprf+'|'+s_noprf
  else
    TFVierge(ecran).Retour := 'mb_ok'+'|'+s_profil;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_PROFIL_GRPTRAVAIL.OnArgument (S : String ) ;
var
  s_user,s_profil : string;
  idx : integer;
begin
  Inherited ;
  AvertirTable('TTUSERGROUPE');
  Obj_combo  := THValComboBox(GetControl('CB_PROFIL'));
  if Obj_combo <> nil then
    Obj_combo.OnChange :=  DoChange;

  Obj_ChkbAllPrf := THCheckbox(GetControl('CHKBX_ALLPROFIL'));
  if Obj_ChkbAllPrf <> nil then
    Obj_ChkbAllPrf.OnClick := DoClick;

  Obj_ChkbNoPrf  := THCheckbox(GetControl('CHKBX_NOPROFIL'));
  if Obj_ChkbNoPrf <> nil then
    Obj_ChkbNoPrf.OnClick := DoClick;

  Obj_BtnErase   := TToolbarButton97(GetControl('BTN_ERASE'));
  if Obj_BtnErase <> nil then
    Obj_BtnErase.OnClick := DoClick;

  // initialisation
  s_user := s;
  s_mode := READTOKENPipe(s,'|');

  if s_mode = 'MODIFICATION' then
  begin
    s_user   := READTOKENPipe(s,'|');
    s_profil := s;
    idx := Obj_combo.Values.IndexOf(s_profil);
    Obj_combo.ItemIndex := idx;

    Obj_ChkbAllPrf.Enabled := FALSE;
    Obj_ChkbNoPrf.Enabled  := FALSE;
  end;
  // intitulé de la fenêtre.
  ecran.Caption := ecran.Caption+' ['+s_user+']';

  if Obj_combo.ItemIndex = -1 then
    Obj_ChkbAllPrf.Enabled := FALSE;
end ;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PROFIL_GRPTRAVAIL.OnClose ;
begin
  Inherited ;
  // Libération
  Obj_combo.free;
  Obj_ChkbAllPrf.free;
  Obj_ChkbNoPrf.free;
  Obj_BtnErase.free;
end ;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PROFIL_GRPTRAVAIL.DoClick(Sender: Tobject);
begin
  // Actions sur le bonton "Effacer" et les cases à cocher.
  if (sender is TToolbarButton97) then
  begin
    // RAZ de la ComboBox.
    Obj_combo.ItemIndex := -1;
    if Obj_ChkbAllPrf.Checked then
      Obj_ChkbAllPrf.Checked := FALSE;
    Obj_ChkbAllPrf.Enabled := FALSE;
  end
  else if (sender is THCheckbox) then
  begin
    if (sender as THCheckbox).Name = Obj_ChkbAllPrf.Name then
    begin
      if Obj_ChkbAllPrf.Checked then
        Obj_ChkbNoPrf.Checked := FALSE;
    end
    else
    begin
      if Obj_ChkbNoPrf.Checked then
      begin
        Obj_ChkbAllPrf.Checked := FALSE;
        Obj_combo.ItemIndex := -1;
      end;
    end;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_PROFIL_GRPTRAVAIL.DoChange(Sender: Tobject);
begin
  //Après selection d'un profil dans la liste déroulante.
  if (Obj_combo.ItemIndex <> -1) and (s_mode <> 'MODIFICATION') then
    Obj_ChkbAllPrf.Enabled := TRUE;
end;
////////////////////////////////////////////////////////////////////////////////
Procedure TOF_PROFIL_GRPTRAVAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PROFIL_GRPTRAVAIL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PROFIL_GRPTRAVAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PROFIL_GRPTRAVAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PROFIL_GRPTRAVAIL.OnLoad ;
begin
  Inherited ;
end ;
////////////////////////////////////////////////////////////////////////////////
Initialization
  registerclasses ( [ TOF_PROFIL_GRPTRAVAIL ] ) ;
end.
