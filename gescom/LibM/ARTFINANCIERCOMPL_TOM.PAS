{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/08/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ARTFINANCIERCOMPL (ARTFINANCIERCOMPL)
Mots clefs ... : TOM;ARTFINANCIERCOMPL
*****************************************************************}
Unit ARTFINANCIERCOMPL_TOM ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,dbtables,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM, Fiche,
     FichList,UTob ;

Type
  TOM_ARTFINANCIERCOMPL = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

  Private
       Journal,compte,CodeArticle : String ;
    end ;

const
	// libellés des messages
	TexteMessage: array[1..4] of string 	= (
          {1}        'Vous devez renseigner un établissement.',
          {2}        'Vous devez renseigner un libellé.',
          {3}        'Vous devez renseigner le journal. ',
          {4}        ''
                     );

Implementation

procedure TOM_ARTFINANCIERCOMPL.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_ARTFINANCIERCOMPL.OnDeleteRecord ;
begin
  Inherited ;
SetControlText('GFC_JOURNAL','') ;
end ;

procedure TOM_ARTFINANCIERCOMPL.OnUpdateRecord ;
begin
  Inherited ;
if GetControlText('GFC_ETABLISSEMENT')='' then
  BEGIN
  SetFocusControl('GFC_ETABLISSEMENT'); LastError:=1; LastErrorMsg:=TexteMessage[LastError]; exit;
  END;
if GetControlText('GFC_LIBELLE')='' then
  BEGIN
  SetFocusControl('GFC_LIBELLE'); LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit;
  END;
if GetControlText('GFC_JOURNAL')='' then //THValComboBox(GetControl('MPC_JALREGLE')).Value='' then
  BEGIN
  SetFocusControl('GFC_JOURNAL'); LastError:=3; LastErrorMsg:=TexteMessage[LastError]; exit;
  END;
end ;

procedure TOM_ARTFINANCIERCOMPL.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_ARTFINANCIERCOMPL.OnLoadRecord ;
begin
  Inherited ;
SetField('GFC_CODEARTICLE',CodeArticle) ;
SetFocusControl('GFC_ETABLISSEMENT');
end ;

procedure TOM_ARTFINANCIERCOMPL.OnChangeField ( F: TField ) ;
begin
  Inherited ;
if (F.FieldName='GFC_ETABLISSEMENT') and (GetControlText('GFC_ETABLISSEMENT')<>'') and (GetControlText('GFC_JALREGLE')='')  then
  begin
 // SetControlText('GFC_JOURNAL',Journal) ;
 // SetControlText('GFC_COMPTE',Compte) ;
  end ;
end ;

procedure TOM_ARTFINANCIERCOMPL.OnArgument ( S: String ) ;
var x:integer ;
critere,ChampMul,ValMul: String ;
begin
Inherited;
Repeat
  Critere:=Trim(ReadTokenSt(S)) ;
  if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
       begin
       ChampMul:=copy(Critere,1,x-1) ;
       ValMul:=copy(Critere,x+1,length(Critere)) ;
       if ChampMul='JOURNAL' then Journal:=ValMul ;
       if ChampMul='COMPTE' then Compte:=ValMul ;
       if ChampMul='CODEARTICLE' then CodeArticle:=ValMul ;
       end ;
    end ;
until Critere='' ;
end;

procedure TOM_ARTFINANCIERCOMPL.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_ARTFINANCIERCOMPL.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_ARTFINANCIERCOMPL ] ) ; 
end.
