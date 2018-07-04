{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 27/04/2004
Modifié le ... :   /  /
Description .. : Gestion de la fiche DECLARANT
Mots clefs ... : PAIE;DECLARANT
*****************************************************************
PT1 09/12/2004 JL V_60 Suppression itemIndex:=0 dans OnArgument
PT2 18/02/2004 SB V_60 FQ 11973 Affichage <<Tous>>
}
unit UTOMDeclarantAttest;

interface

uses
  Classes,
  SysUtils,
  Utom,
{$IFNDEF EAGLCLIENT}
  db,HDB,
{$ELSE}
  Utob,MainEAGL,EFiche,
{$ENDIF}
  PgOutils,
  HmsgBox,
  HCtrls
  ;

Type
    TOM_DECLARANTATTEST = class(TOM)
    procedure OnChangeField(F: TField);        override;
    procedure OnUpdateRecord;                  override;
    procedure OnArgument ( Argument : String); override;
    procedure OnLoadRecord;                    override;
    End;

implementation

{ TOM_DECLARANTATTEST }

procedure TOM_DECLARANTATTEST.OnArgument(Argument: String);
begin
  inherited;
end;

procedure TOM_DECLARANTATTEST.OnChangeField(F: TField);
begin
  inherited;
  If (F.FieldName=('PDA_DECLARANTATTES')) AND (DS.State in [DsInsert]) then
    Begin
    If GetField('PDA_DECLARANTATTES') <> UpperCase(GetField('PDA_DECLARANTATTES')) then
       SetField('PDA_DECLARANTATTES', UpperCase(GetField('PDA_DECLARANTATTES')));
    End;

  If F.FieldName=('PDA_QUALDECLARANT') then
    Begin
    If GetField('PDA_QUALDECLARANT') = 'AUT' then
       SetControlEnabled('PDA_AUTRE',True)
    else
      Begin
      SetControlEnabled('PDA_AUTRE',False);
      If GetField('PDA_AUTRE')<>'' then SetField('PDA_AUTRE','');
      End;
    End;

  If F.FieldName=('PDA_CIVILITE') then
    Begin
    If ((GetField('PDA_CIVILITE')='MME') OR (GetField('PDA_CIVILITE')='MLE')) AND (GetField('PDA_SEXE')<>'F') then
      SetField('PDA_SEXE','F')
    else
      If (GetField('PDA_CIVILITE')='MR') AND (GetField('PDA_SEXE')<>'M') then
        SetField('PDA_SEXE','M');
    End;

  If F.FieldName=('PDA_SEXE') then
    Begin
      If (GetField('PDA_SEXE')='M') AND (GetField('PDA_CIVILITE')<>'MR')  then
        SetField('PDA_CIVILITE','MR');
    End;

end;


procedure TOM_DECLARANTATTEST.OnLoadRecord;
begin
{ DEB PT2 }
  if GetControlText('PDA_ETABLISSEMENT')='' then
     SetControlProperty('PDA_ETABLISSEMENT','Text','<<Tous>>');
  if GetControlText('PDA_TYPEATTEST')='' then
     SetControlProperty('PDA_TYPEATTEST','Text','<<Tous>>');
{ FIN PT2 }
end;

procedure TOM_DECLARANTATTEST.OnUpdateRecord;
var
{$IFNDEF EAGLCLIENT}
MultiEtab,MultiType : THDBMultiValComboBox;
{$ELSE}
MultiEtab,MultiType : THMultiValComboBox;
{$ENDIF}
begin
  inherited;
  If (GetField('PDA_DECLARANTATTES') = '' ) Or (GetField('PDA_LIBELLE') = '' ) then exit;
{  If (GetField('PDA_ETABLISSEMENT') = '' ) then
    Begin
    If ExisteSql('SELECT PDA_ETABLISSEMENT FROM DECLARANTATTEST '+
    ' WHERE PDA_ETABLISSEMENT="" AND PDA_DECLARANTATTES<>"'+GetField('PDA_DECLARANTATTES')+'" ') then
       Begin
       LastError := 1 ;
       PgiBox('Vous ne pouvez définir deux déclarants pour tous les établissements.',Ecran.Caption);
       Exit;
       End;
    End;   }

  If Getfield('PDA_QUALDECLARANT') = '' then
    Begin
    LastError := 1 ;
    PgiBox('Veuiller renseigner la qualité du déclarant.',Ecran.Caption);
    Exit;
    End;

  If Getfield('PDA_SEXE') = '' then
    Begin
    LastError := 1 ;
    PgiBox('Veuiller renseigner le sexe du déclarant.',Ecran.Caption);
    Exit;
    End;
  {$IFNDEF EAGLCLIENT}
  MultiEtab := THDBMultiValComboBox(GetControl('PDA_ETABLISSEMENT'));
  MultiType := THDBMultiValComboBox(GetControl('PDA_TYPEATTEST'));
  {$ELSE}
  MultiEtab := THMultiValComboBox(GetControl('PDA_ETABLISSEMENT'));
  MultiType := THMultiValComboBox(GetControl('PDA_TYPEATTEST'));
  {$ENDIF}
  If MultiEtab.Tous = True then SetField('PDA_ETABLISSEMENT','');
  If MultiType.Tous = True then SetField('PDA_TYPEATTEST','');
end;

initialization
  registerclasses([TOM_DECLARANTATTEST]);
end.
