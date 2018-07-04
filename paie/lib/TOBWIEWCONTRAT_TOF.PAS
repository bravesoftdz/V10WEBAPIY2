{***********UNITE*************************************************
Auteur  ...... : Gérard GUILLAUD-SAUMUR
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TOBWIEWCONTRAT ()
Mots clefs ... : TOF;TOBWIEWCONTRAT
*****************************************************************}
Unit TOBWIEWCONTRAT_TOF ;

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
     UTOF ; 

Type
  TOF_TOBWIEWCONTRAT = Class (TOF)
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure CkarreteClick (Sender : TObject);
  end ;

Implementation

procedure TOF_TOBWIEWCONTRAT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_TOBWIEWCONTRAT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_TOBWIEWCONTRAT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_TOBWIEWCONTRAT.OnLoad ;
var
  Datearret : TDateTime;
  where : String;
begin
  Inherited ;
  if TCheckBox(GetControl('ARRETE_')) <> nil then
  begin
    if (GetControlText('arrete_')='X') and (IsValidDate(GetControlText('ARRETE'))) then
    begin
      datearret := StrToDate(GetControlText('arrete'));
      where := ' PCI_DEBUTCONTRAT <= "'+ USDateTime(datearret) +'" AND (PCI_FINCONTRAT >= "'
              + USDateTime(datearret) + '" OR PCI_FINCONTRAT= "' + USDateTime(iDATE1900) + '" OR PCI_FINCONTRAT IS NULL)';
      SetControlText ('XX_WHERE',where)
    end;
  end;
end ;

procedure TOF_TOBWIEWCONTRAT.OnArgument (S : String ) ;
var arrete : TcheckBox;
  Edit : THEdit;
begin
  Inherited ;
  arrete := TCheckBox(GetControl('ARRETE_'));
  If arrete <> Nil Then arrete.OnCLick := ckarreteClick;
  Edit := THEdit(GetControl('ARRETE'));
end ;

procedure TOF_TOBWIEWCONTRAT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_TOBWIEWCONTRAT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_TOBWIEWCONTRAT.OnCancel () ;
begin
  Inherited ;
end ;
procedure TOF_TOBWIEWCONTRAT.CkarreteClick(Sender : TObject);
var STWhere : String;
begin
        If Sender = Nil then Exit;
        If GetCheckBoxState('ARRETE_') = CbChecked then
        begin
                SetControlEnabled('ARRETE',True);
                SetControlEnabled('LABELDATE',True);
                SetControlText('ARRETE',DateToStr(Date));
                StWhere := ' PCI_DEBUTCONTRAT <= "'+ USDateTime(date) +'" AND (PCI_FINCONTRAT >= "'
              + USDateTime(date) + '" OR PCI_FINCONTRAT= "' + USDateTime(iDATE1900) + '" OR PCI_FINCONTRAT IS NULL)';
        end
        else
        begin
                SetControlEnabled('ARRETE',False);
                SetControlEnabled('LABELDATE',False);
                StWhere := '';
        end;
        SetControlText('XX_WHERE',StWhere);
end;

Initialization
  registerclasses ( [ TOF_TOBWIEWCONTRAT ] ) ; 
end.
