{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 31/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPCONTROLTVA ()
Mots clefs ... : TOF;CPCONTROLTVA
*****************************************************************}
Unit CPCONTROLTVA_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     Fe_main,
     Qrs1,
{$else}
     eMul,
     uTob,
     Maineagl,
     eQrs1,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     ParamSoc,
     SaisUtil , // pour le CaseNatJal
     Ent1 , // pour le tzJvente
     Dialogs,
     UTOF ;

Type
  TOF_CPCONTROLTVA = Class (TOF)
   private
    procedure TYPETVAChange( Sender : TObject ) ;
    procedure FETATChange( Sender : TObject ) ;
   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;


function CPLanceFiche_CPCONTROLTVA( vStTypeTva : string = '' ) : string;

Implementation

function CPLanceFiche_CPCONTROLTVA( vStTypeTva : string = '' ) : string;
begin
  result := AGLLanceFiche('CP','CPECONTROLTVA', '', '', vStTypeTva);
end;

procedure TOF_CPCONTROLTVA.OnNew ;
begin
 Inherited ;
end ;

procedure TOF_CPCONTROLTVA.OnDelete ;
begin
 Inherited ;
end ;

procedure TOF_CPCONTROLTVA.OnUpdate ;
begin
 Inherited ;
end ;

procedure TOF_CPCONTROLTVA.OnLoad ;
begin
 Inherited ;
 if THValCombobox(GetControl('FETAT' , true )).Value = 'EC1' then
  SetControlText('XX_WHERE', ' AND Y_AXE="' + GetParamSocSecur('SO_CPPCLAXETVA', '  ') + '" ' )
   else
    SetControlText('XX_WHERE', '' ) ;
end ;

procedure TOF_CPCONTROLTVA.OnArgument ( S : String ) ;
var lSt : string;
begin
 Inherited ;
 THMultiValCombobox(GetControl('TYPETVA', true)).OnChange := TYPETVAChange ;
 THMultiValCombobox(GetControl('FETAT', true)).OnChange := FETATChange ;

 // GCO - 26/07/2007 - FQ 20184
 if S <> '' then
 begin
   TFQRS1(Ecran).CodeEtat := 'EC2';
   THMultiValCombobox(GetControl('TYPETVA', True)).Value:= S;
   TYPETVAChange( nil ) ;
   THEdit(GetControl('E_DATECOMPTABLE', True)).Text  := DateToStr(VH^.EnCours.Deb);
   THEdit(GetControl('E_DATECOMPTABLE_', True)).Text := DateToStr(VH^.EnCours.Fin);
 end;

 lSt := TraduireMemoire('<<Tous>>');
 THMultiValCombobox(GetControl('J_JOURNAL', True)).SelectAll;
 if THMultiValCombobox(GetControl('J_JOURNAL', True)).Tous then
   THMultiValCombobox(GetControl('J_JOURNAL', True)).text := lSt;

 lSt := TraduireMemoire('<<Tous>>');
 THMultiValCombobox(GetControl('TYPETVA', True)).SelectAll;
 if THMultiValCombobox(GetControl('TYPETVA', True)).Tous then
   THMultiValCombobox(GetControl('TYPETVA', True)).text := lSt;
end ;

procedure TOF_CPCONTROLTVA.OnClose ;
begin
 Inherited ;
end ;

procedure TOF_CPCONTROLTVA.OnDisplay () ;
begin
 Inherited ;
end ;

procedure TOF_CPCONTROLTVA.OnCancel () ;
begin
 Inherited ;
end ;

procedure TOF_CPCONTROLTVA.TYPETVAChange( Sender : TObject ) ;
begin
 { FQ 19740 & 19195 - CA - 31/05/2007 - On ne doit pas afficher les journaux d'OD }
 if GetControlText('TYPETVA') = 'COLL;' then
 begin
  THMultiValCombobox(GetControl('J_JOURNAL', true)).Plus := ' J_FERME="-" AND (J_NATUREJAL="VTE" OR J_NATUREJAL="BQE" OR J_NATUREJAL="CAI") ';
  SetControlText('NATUREGENE','PRO');
 end
 else if GetControlText('TYPETVA') = 'DED;' then
 begin
  THMultiValCombobox(GetControl('J_JOURNAL', true)).Plus := ' J_FERME="-" AND (J_NATUREJAL="ACH" OR J_NATUREJAL="BQE" OR J_NATUREJAL="CAI") ';
  SetControlText('NATUREGENE','CHA","IMO');
 end
 else
 begin
  THMultiValCombobox(GetControl('J_JOURNAL', true)).Plus := ' J_FERME="-" AND (J_NATUREJAL="ACH" OR J_NATUREJAL="VTE"  OR J_NATUREJAL="BQE" OR J_NATUREJAL="CAI") ' ;
  SetControlText('NATUREGENE','PRO","CHA","IMO');
 end;

 THMultiValCombobox(GetControl('J_JOURNAL', true)).Reload ;

end ;

procedure TOF_CPCONTROLTVA.FETATChange(Sender: TObject);
begin
  SetControlVisible ('TYPETVA',GetControlText('FEtat')<>'EC1');
  SetControlVisible ('LTYPETVA',GetControlText('FEtat')<>'EC1');
end;

Initialization
  registerclasses ( [ TOF_CPCONTROLTVA ] ) ;
end.
