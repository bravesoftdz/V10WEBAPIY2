{***********UNITE*************************************************
Auteur  ...... : GGR
Créé le ...... : 26/04/2006
Modifié le ... :   /  /
Description .. : Liste des anomalies d'affectation de salariés aux comptes auxiliaires
Mots clefs ... : TOF;ANOAFFECTAUX_ETAT
*****************************************************************}
Unit ANOAFFECTAUX_ETAT_TOF ;
{
PT1 : 30/08/2007 PH V_80 FQ N°14714 Ergonomie
}
Interface

Uses StdCtrls, Controls, Classes,Graphics,forms,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,EdtEtat,QRS1,
{$else}
     eMul,
     uTob,
     eQrs1,
{$ENDIF}
     sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, PgEditOutils,PgEditOutils2,
     HTB97,HPdfviewer,ParamDat,EntPaie, HQry, PGoutils, PGoutils2;

Type
  TOF_ANOAFFECTAUX_ETAT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure OnClickSalarieSortie(Sender: TObject);
  end ;

Implementation

procedure TOF_ANOAFFECTAUX_ETAT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ANOAFFECTAUX_ETAT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ANOAFFECTAUX_ETAT.OnUpdate ;
var
Check : TCheckBox;
regl :  THMultiValComboBox;
Where, sval, stval, s1where, s2where, Swhere : String;
Pages : TPageControl;
begin
  Inherited ;
  Pages := TPageControl(GetControl('Pages'));
  Where := RecupWhereCritere(Pages);
  Check:=TCheckBox(GetControl('CKSORTIE'));
  regl  := THMultiValComboBox(GetControl('PSA_PGMODEREGLE'));
  if Check=nil then
    Begin
      SetControlVisible('DATEARRETE',False);
      SetControlVisible('TDATEARRETE',False);
    End
  else
    Check.OnClick:=OnClickSalarieSortie;
  if not regl.tous then
    begin
      sval := regl.text ;
      s1where := '' ;
      s2where := '' ;
      stval := '';
      while sval <> '' do
      begin
        Stval := ReadTokenPipe(sval,';');
        if s1where = '' then s1where := 'OR (PSA_PAIACOMPTE="'+Stval+'"'
        else s1where := 'OR PSA_PAIACOMPTE="'+Stval+'"';
        if s2where = '' then s2where := ' OR (PSA_PAIFRAIS="'+Stval+'"'
        else s2where := 'OR PSA_PAIFRAIS="'+Stval+'"';
      end;
    end;
  if s1where <> '' then s1where := s1where + ')';
  if s2where <> '' then s2where := s2where + ')';
  Swhere := s1where + s2where;
  WHERE := WHERE + Swhere;
  SetControlText('XX_WHERE',WHERE);
end ;

procedure TOF_ANOAFFECTAUX_ETAT.OnLoad ;
var
  DateArret : TDateTime;
  St        : String;
begin
inherited ;
if  TCheckBox(GetControl('CKSORTIE'))<>nil then
  Begin
   SetControlText('XX_WHERE','');
   if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRETE')))then
     Begin
     DateArret:=StrToDate(GetControlText('DATEARRETE'));
     st:=' (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'"'+
         ' OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'"'+
         ' OR PSA_DATESORTIE IS NULL)'+
         ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'" ';
     SetControlText('XX_WHERE',GetControlText('XX_WHERE') + st);
     End;
  End;
end;

procedure TOF_ANOAFFECTAUX_ETAT.OnClickSalarieSortie(Sender: TObject);
begin
  SetControlenabled('DATEARRETE',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRETE',(GetControltext('CKSORTIE')='X'));
//  SetControlText('DATEARRETE',''); // PT1
end;

procedure TOF_ANOAFFECTAUX_ETAT.OnArgument (S : String ) ;
var
Check : TCheckBox;
begin
  Inherited ;
      TFQRS1(Ecran).Caption := 'Liste d''Anomalies d''affectation aux comptes auxiliaires';
      TFQRS1(Ecran).CodeEtat:= 'PAX';
  SetControlvisible('DATEARRETE',True);
  SetControlvisible('TDATEARRETE',True);
  SetControlEnabled('DATEARRETE',False);
  SetControlEnabled('TDATEARRETE',False);
  Check:=TCheckBox(GetControl('CKSORTIE'));
  if Check=nil then
    Begin
      SetControlVisible('DATEARRETE',False);
      SetControlVisible('TDATEARRETE',False);
    End
  else
    Check.OnClick:=OnClickSalarieSortie;
end ;

procedure TOF_ANOAFFECTAUX_ETAT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_ANOAFFECTAUX_ETAT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_ANOAFFECTAUX_ETAT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_ANOAFFECTAUX_ETAT ] ) ;
end.
