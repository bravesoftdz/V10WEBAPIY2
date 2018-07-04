{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/01/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AUGMEXCLUS (AUGMEXCLUS)
Mots clefs ... : TOM;AUGMEXCLUS
*****************************************************************}
Unit UTomAugmExclus ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM,
     Hqry, 
     UTob ;

Type
  TOM_AUGMEXCLUS = Class (TOM)
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
    private
    LeCritere : String;
    procedure ChangeValeur(Sender : Tobject);
    end ;

Implementation

procedure TOM_AUGMEXCLUS.OnNewRecord ;
begin
  Inherited ;
  SetField('PAE_ORDRE',-1);
  SetField('PAE_CRITEREAUGM',LeCritere);
end ;

procedure TOM_AUGMEXCLUS.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_AUGMEXCLUS.OnUpdateRecord ;
var NumOrdre : Integer;
    Q : TQuery;
    Combo : THValComboBox;
    Edit : THEdit;
begin
  Inherited ;
  If GetField('PAE_CRITEREAUGEX') = '' then
  begin
       PGIBox('Vous devez renseigner le critère',Ecran.Caption);
       LastError :=1;
       Exit;
  end;
  If GetField('PAE_OPERATTEST') = '' then
  begin
       PGIBox('Vous devez renseigner l''opérateur',Ecran.Caption);
       LastError :=1;
       Exit;
  end;
  If GetField('PAE_ORDRE') = -1 then
  begin
       Q := OpenSQL('SELECT MAX(PAE_ORDRE) NUMERO FROM AUGMEXCLUS WHERE PAE_CRITEREAUGM="'+Getfield('PAE_CRITEREAUGM')+'"',True);
       If not Q.Eof then NumOrdre := Q.FindField('NUMERO').AsInteger
       else NumOrdre := 0;
       Ferme(Q);
       SetField('PAE_ORDRE',NumOrdre + 1);
  end;
  Combo := THValComboBox(GetControl('ETABLETTE'));
  Edit := THEdit(GetControl('EDATA'));
  If Combo.Visible = true then SetField('PAE_VALEUREX',GetControltext('ETABLETTE'))
  else If Edit.Visible = true then SetField('PAE_VALEUREX',GetControlText('EDATA'))
  else SetField('PAE_VALEUREX','');
  SetField('PAE_LIBELLE',RechDom('PGAUGMCRITERES',GetField('PAE_CRITEREAUGEX'),False));
end ;

procedure TOM_AUGMEXCLUS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_AUGMEXCLUS.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_AUGMEXCLUS.OnLoadRecord ;
var Q : TQuery;
    TypeValeur : String;
begin
  Inherited ;
  Q := OpenSQL('SELECT CO_ABREGE FROM COMMUN WHERE CO_CODE="'+GetField('PAE_CRITEREAUGEX')+'" AND CO_TYPE="PGM"',True);
  If Not Q.Eof then TypeValeur := Q.FindField('CO_ABREGE').AsString
  else TypeValeur := 'DATA';
  Ferme(Q);
  If TypeValeur = 'DATA' then
  begin
       SetControlVisible('EDATA',True);
       SetControlVisible('ETABLETTE',False);
       SetControlText('EDATA',GetField('PAE_VALEUREX'));
  end
  else
  begin
       SetControlVisible('EDATA',False);
       SetControlVisible('ETABLETTE',True);
       SetControlProperty('ETABLETTE','Datatype',TypeValeur);
       SetControlText('ETABLETTE',GetField('PAE_VALEUREX'));
  end;
end ;

procedure TOM_AUGMEXCLUS.OnChangeField ( F: TField ) ;
var TypeValeur : String;
    Q : TQuery;
begin
  Inherited ;
  If F.FieldName = 'PAE_CRITEREAUGEX' then
  begin
       If GetField('PAE_CRITEREAUGEX') <> '' then
       begin
              Q := OpenSQL('SELECT CO_ABREGE FROM COMMUN WHERE CO_CODE="'+GetField('PAE_CRITEREAUGEX')+'" AND CO_TYPE="PGM"',True);
              If Not Q.Eof then TypeValeur := Q.FindField('CO_ABREGE').AsString
              else TypeValeur := 'DATA';
              Ferme(Q);
              If TypeValeur = 'DATA' then
              begin
                   SetControlVisible('EDATA',True);
                   SetControlVisible('ETABLETTE',False);
              end
              else
              begin
                   SetControlVisible('EDATA',False);
                   SetControlVisible('ETABLETTE',True);
                   SetControlProperty('ETABLETTE','Datatype',TypeValeur);
              end;
       end
       else
       begin
            SetControlVisible('EDATA',False);
            SetControlVisible('ETABLETTE',False);
            SetControltext('EDATA','');
            SetControltext('ETABLETTE','');
       end;
  end;
end ;

procedure TOM_AUGMEXCLUS.OnArgument ( S: String ) ;
var Combo : THValComboBox;
    Edit : THEdit;
begin
  Inherited ;
  LeCritere := ReadTokenPipe(S,';');
  Combo := THValComboBox(GetControl('ETABLETTE'));
  Edit := THEdit(GetControl('EDATA'));
  If Combo <> Nil then Combo.OnChange := ChangeValeur;
  If Edit <> Nil then Edit.OnChange := ChangeValeur;
end ;

procedure TOM_AUGMEXCLUS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_AUGMEXCLUS.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_AUGMEXCLUS.ChangeValeur(Sender : Tobject);
var St : String;
begin
     If GetControlVisible('EDATA') = True then St := GetControlText('EDATA');
     If GetControlVisible('ETABLSETTE') = True then St := GetControlText('ETABLETTE');
     If GetField('PAE_VALEUREX') <> St then ForceUpdate;
end;

Initialization
  registerclasses ( [ TOM_AUGMEXCLUS ] ) ;
end.


