{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULGENE ()
Mots clefs ... : TOF;MULGENE
*****************************************************************}
Unit MULGENE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$ifdef eAGLClient}
  MenuOLX,MaineAGL,eTablette,
{$else}
  Tablette,FE_Main,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
{$endif eAGLClient}

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
     Ent1,
     Mul ;

Type
  TOF_MULGENE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure VentilOnChange(Sender: TObject);
    procedure FListeDblClick(sender: TObject);
    procedure VentilS3OnClick(sender: TObject);
  end ;

Implementation

procedure TOF_MULGENE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULGENE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULGENE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULGENE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULGENE.VentilOnChange (Sender : TObject ) ;
var ValeurVentil : string ;
begin
  ValeurVentil := GetControlText('VENTILABLE') ;
  if ValeurVentil = '' then SetControlText('XX_WHERE', '' ) ;
  if ValeurVentil = 'A1' then SetControlText('XX_WHERE', 'G_VENTILABLE1="X"' ) ;
  if ValeurVentil = 'A2' then SetControlText('XX_WHERE', 'G_VENTILABLE2="X"' ) ;
  if ValeurVentil = 'A3' then SetControlText('XX_WHERE', 'G_VENTILABLE3="X"' ) ;
  if ValeurVentil = 'A4' then SetControlText('XX_WHERE', 'G_VENTILABLE4="X"' ) ;
  if ValeurVentil = 'A5' then SetControlText('XX_WHERE', 'G_VENTILABLE5="X"' ) ;
  if ValeurVentil = 'NON' then SetControlText('XX_WHERE', 'G_VENTILABLE="-"' ) ;
  if ValeurVentil = 'QUE' then SetControlText('XX_WHERE', 'G_VENTILABLE="X"' ) ;
end ;

procedure TOF_MULGENE.VentilS3OnClick ( sender : TObject ) ;
begin
  case TCheckbox(GetControl('VENTILS3')).State of
    cbchecked   : THValComboBox(GetControl('VENTILABLE')).Value := 'A1' ;
    cbunchecked : THValComboBox(GetControl('VENTILABLE')).Value := 'NON' ;
    cbgrayed    : THValComboBox(GetControl('VENTILABLE')).Value := '' ;
  end ;  
end ;

procedure TOF_MULGENE.FListeDblClick (sender : TObject ) ;
begin
  AGLLanceFiche('CP','CPGENERAL','' ,TFMul(Ecran).Q.FindField('G_GENERAL').AsString ,'' ) ;
end ;

procedure TOF_MULGENE.OnArgument (S : String ) ;
begin
  Inherited ;
  SetControlText('G_DATEMODIF', StDate1900 ) ;
  SetControlText('G_DATEMODIF_', StDate2099 ) ;
  SetControlText('G_DATEDERNMVT', StDate1900 ) ;
  SetControlText('G_DATEDERNMVT_', StDate2099 ) ;
  LibellesTableLibre(TTabSheet(GetControl('PZLIBRE')),'TG_TABLE','G_TABLE','G') ;
  THValComboBox(GetControl('VENTILABLE')).onChange := VentilOnChange ;
  TFMul(Ecran).FListe.OnDblClick := FListeDblClick ;
  if EstSerie(S3) then
    begin
    SetControlVisible('VENTILABLE', false ) ;
    SetControlVisible('TVENTILABLE', false ) ;
    SetControlVisible('VENTILS3', true ) ;
    end ;
  TCheckBox(GetControl('VENTILS3')).OnClick := VentilS3OnClick ;
end ;

procedure TOF_MULGENE.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MULGENE ] ) ;
end.
