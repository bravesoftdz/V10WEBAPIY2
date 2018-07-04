{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPJALDIV ()
Mots clefs ... : TOF;CPJALDIV
*****************************************************************}
Unit CPJALDIV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
{$IFDEF EAGLCLIENT}
     eQRS1,MaineAGL,
{$ELSE}
     QRS1,FE_Main,
{$ENDIF}
      Ent1;

Type
  TOF_CPJALDIV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure NormalOnClick(sender: TObject);
    procedure AffichageOnClick(sender: TObject);
    procedure SimulOnClick(sender: TObject);
    procedure SituationOnClick(sender: TObject);
    procedure ExerciceOnChange(sender: TObject);
    procedure TriDateOnClick(sender: TObject);
    procedure IncJalAnoOnClick(sender: TObject);
  end ;

Implementation

procedure TOF_CPJALDIV.NormalOnClick (sender : TObject ) ;
begin
  if GetCheckBoxState('NORMAL')=cbChecked then SetControlText('N', 'N') else SetControlText('N', 'Z' ) ;
end ;

procedure TOF_CPJALDIV.SimulOnClick (sender : TObject ) ;
begin
  if GetCheckBoxState('SIMUL')=cbChecked then SetControlText('S', 'S') else SetControlText('S', 'Z' ) ;
end ;

procedure TOF_CPJALDIV.SituationOnClick (sender : TObject ) ;
begin
  if GetCheckBoxState('SITUATION')=cbChecked then SetControlText('U', 'U') else SetControlText('U', 'Z' ) ;
end ;

procedure TOF_CPJALDIV.AffichageOnClick (sender : TObject ) ;
begin
  Case THRadioGroup(GetControl('AFFICHAGE')).ItemIndex of
    0 : begin
        SetControlText('DEBIT', 'E_DEBIT' ) ;
        SetControlText('CREDIT', 'E_CREDIT' ) ;
        end ;
    1 : begin
        SetControlText('DEBIT', 'E_DEBITDEV' ) ;
        SetControlText('CREDIT', 'E_CREDITDEV' ) ;
        end ;
    2 : begin
        SetControlText('DEBIT', 'E_DEBITEURO' ) ;
        SetControlText('CREDIT', 'E_CREDITEURO' ) ;
        end ;
  end ;
  SetControlText('LIBAFFDEV', THRadioGroup(GetControl('AFFICHAGE')).Items[THRadioGroup(GetControl('AFFICHAGE')).ItemIndex] ) ;
  SetControlText('SMON', THRadioGroup(GetControl('AFFICHAGE')).Values[THRadioGroup(GetControl('AFFICHAGE')).ItemIndex] ) ;
end ;

procedure TOF_CPJALDIV.TriDateOnClick (sender : TObject ) ;
begin
  SetControlText('XX_ORDERBY', 'E_JOURNAL' ) ;
  if GetCheckBoxState('TRIDATE')=cbChecked then SetControlText('XX_ORDERBY', GetControlText('XX_ORDERBY')+',E_EXERCICE,E_DATECOMPTABLE' ) ;
  SetControlText('XX_ORDERBY', GetControlText('XX_ORDERBY')+',E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE' ) ;
end ;

procedure TOF_CPJALDIV.IncJalAnoOnClick (sender : TObject ) ;
begin
  SetControlText('XX_WHERE', 'J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA"') ;
  if GetCheckBoxState('INCJALANO')=cbUnChecked then SetControlText('XX_WHERE', GetControlText('XX_WHERE')+' AND J_NATUREJAL<>"ANO" ') ;
end ;

procedure TOF_CPJALDIV.ExerciceOnChange (sender : TObject) ;
begin
  ExoToDates(GetControlText('E_EXERCICE'),GetControl('E_DATECOMPTABLE'),GetControl('E_DATECOMPTABLE_')) ;
end ;

procedure TOF_CPJALDIV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPJALDIV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPJALDIV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPJALDIV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPJALDIV.OnArgument (S : String ) ;
begin
  Inherited ;
  // gestion de la selection du type d'écriture
  if GetControlText('NORMAL') ='X' then SetControlText('N','N') else SetControlText('N','Z');
  if GetControlText('SIMUL') ='X' then SetControlText('S','S') else SetControlText('S','Z');
  if GetControlText('SITUATION') ='X' then SetControlText('U','U') else SetControlText('U','Z');
  TCheckBox(GetControl('NORMAL')).OnClick := NormalOnClick ;
  TCheckBox(GetControl('SIMUL')).OnClick := SimulOnClick ;
  TCheckBox(GetControl('SITUATION')).OnClick := SituationOnClick ;

  // gestion affichage des montants
  THRadioGroup(GetControl('AFFICHAGE')).Items[0] := VH^.LibDevisePivot ;
  THRadioGroup(GetControl('AFFICHAGE')).Items[2] := VH^.LibDeviseFongible ;
  THRadioGroup(GetControl('AFFICHAGE')).Values[0] := V_PGI.DevisePivot ;
  THRadioGroup(GetControl('AFFICHAGE')).Values[2] := V_PGI.DeviseFongible ;
  THRadioGroup(GetControl('AFFICHAGE')).OnClick := AffichageOnClick ;
  THRadioGroup(GetControl('AFFICHAGE')).itemindex:=0;
  SetControlText('LIBAFFDEV', VH^.LibDevisePivot ) ;
  SetControlText('SMON', V_PGI.DevisePivot ) ;
  SetControlText('DEBIT','E_DEBIT');
  SetControlText('CREDIT','E_CREDIT');

  TCheckBox(GetControl('TRIDATE')).OnClick := TriDateOnClick ;
  TCheckBox(GetControl('INCJALANO')).onClick := IncJalAnoOnClick ;

  THValComboBox(GetControl('E_ETABLISSEMENT')).ItemIndex := 0 ;
  THValComboBox(GetControl('E_DEVISE')).ItemIndex := 0 ;

  SetControlText('E_EXERCICE', VH^.Encours.Code ) ;
  ExoToDates(GetControlText('E_EXERCICE'),GetControl('E_DATECOMPTABLE'),GetControl('E_DATECOMPTABLE_')) ;
  THValComboBox(GetControl('E_EXERCICE')).OnChange := ExerciceOnChange ;

end ;

procedure TOF_CPJALDIV.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPJALDIV ] ) ; 
end.
