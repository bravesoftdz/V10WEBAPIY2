{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : RHCOMPETENCES (RHCOMPETENCES)
Mots clefs ... : TOM;RHCOMPETENCES
*****************************************************************
PT1 24/11/2004 V_60 JL Modifs pour GPAO
}
Unit RHCOMPETENCES_TOM ;

Interface

Uses Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
{$else}
     eFiche,
     eFichList,
     uTob,
{$ENDIF}   
     sysutils,
     HCtrls,
     HMsgBox,
  UTOM,
  HEnt1;

Type
  TOM_RHCOMPETENCES = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation
const
  TexteMessage: array[1..2] of string = (
        {1}  'Vous ne pouvez supprimer cette compétence car des ressources l''utilisent.',
        {2}  'Vous ne pouvez supprimer cette compétence car des compétences par article l''utilisent.'
//mcd 28/11/2005 12581 GA        {2}  'Suppression impossible. Cette compétence est utilisée dans les compétences par article.'
  );

procedure TOM_RHCOMPETENCES.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETENCES.OnDeleteRecord ;
begin
  Inherited ;
  if ExisteSQL ('SELECT PCH_COMPETENCE FROM RHCOMPETRESSOURCE WHERE PCH_COMPETENCE="' + GetField ('PCO_COMPETENCE') + '"') then
  begin
    LastError := 1;
    PGIBox (TraduireMemoire(TexteMessage[1]), Ecran.Caption);
    Exit;
  end;
  if ((ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte)) //AB-200509-FQ 12449
  and ExisteSQL ('SELECT AAC_ARTICLE FROM AFCOMPETARTICLE WHERE AAC_COMPETENCE="' + GetField ('PCO_COMPETENCE') + '"') then
  begin
    LastError := 1;
    PGIBox (TraduireMemoire(TexteMessage[2]), Ecran.Caption);
    Exit;
  end;
end ;

procedure TOM_RHCOMPETENCES.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETENCES.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETENCES.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETENCES.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETENCES.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETENCES.OnArgument ( S: String ) ;
var
  i   : Integer;
  st  : string;
  FieldName : String;
Begin
Inherited;
{-- Libellés libres ------------------------------------------------------------------------------}
  for i := 1 to 5 do
  begin
    FieldName := 'PCO_TABLELIBRERH'+IntToStr(I);
    if assigned(getControl(FieldName)) then
    begin
      St := RechDom('GCZONELIBRE','RH'+IntToStr(I),FALSE);
      SetControlText('T'+FieldName, St);
      SetControlProperty('T' + FieldName, 'VISIBLE', GetControlText('T' + FieldName) <> '.-');
      SetControlProperty(FieldName, 'VISIBLE', GetControlText('T' + FieldName) <> '.-');
    end;
  end;
{-------------------------------------------------------------------------------------------------}
  {$IFDEF SAV}
  Ecran.helpContext := 114000201;
  {$ENDIF SAV}
End;

procedure TOM_RHCOMPETENCES.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_RHCOMPETENCES.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_RHCOMPETENCES ] ) ; 
end.
