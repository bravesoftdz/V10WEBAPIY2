{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMSOCS ()
Mots clefs ... : TOF;BTPARAMSOCS
*****************************************************************}
Unit BTPARAMSOCS ;

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
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     UtilPGI,
     HMsgBox,
     Pays,
     HTB97,
     UtilLine,
     BTPARAMSOCGEN_TOF,
     UTOF ;

Type
  TOF_BTPARAMSOCS = Class (TOF_BTPARAMSOCGEN)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    BEquerrePays : TToolbarButton97;
  	function VerifInfos : boolean;
    procedure BEquerrePaysClick(Sender : TOBJect);
  end ;

Implementation
uses CPTypeCons;

procedure TOF_BTPARAMSOCS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCS.OnUpdate ;
begin
  Inherited ;

  if not verifInfos then
  	 begin
  	 TForm(Ecran).ModalResult:=0;
     exit;
     end;

	StockeInfos(ecran, LaTob);

end ;

procedure TOF_BTPARAMSOCS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCS.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  if GetControl('BEQUERREPAYS') <> nil then
     begin
     BEquerrePays := TToolbarButton97(ecran.FindComponent('BEQUERREPAYS'));
     BEquerrePays.onclick := BEquerrePaysClick;
     end;
  //
  chargeEcran(ecran, LaTob);
end ;

procedure TOF_BTPARAMSOCS.BEquerrePaysClick(Sender : TOBJect);
Begin
 OuvrePays;
End;

procedure TOF_BTPARAMSOCS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMSOCS.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTPARAMSOCS.VerifInfos: boolean;
var CH : Tedit;
begin

	result := false;
  CH:=TEdit(GetCOntrol('SO_LIBELLE')) ;
  if CH.text='' then
  	 Begin
  	 HShowMessage('1;Société;Vous devez renseigner un libellé.;W;O;O;O;','','') ;
     CH.SetFocus ;
     Exit ;
  	 End;

  CH:=TEdit(GetCOntrol('SO_ADRESSE1')) ;
  if CH.Text='' then
  	 Begin
  	 HShowMessage('3;Société;Vous devez renseigner une adresse.;W;O;O;O;','','') ;
     CH.SetFocus ;
     Exit ;
  	 End;

  if (VH^.PaysLocalisation=CodeISOES) and  (CodeISOduPays(GetControlText('SO_PAYS'))=CodeISOES) then
  	 Begin
     CH:=TEdit(GetControl('SO_SIRET')) ;
     if CH.Text <> VerifNIF_ES(CH.Text) then
        BEGIN
        HShowMessage('3;Société;Le NIF/CIF saisie est incorrecte.;W;O;O;O;','','') ;
        CH.SetFocus ;
        Exit ;
        End ;
     End;

  result := true;

end;

Initialization
  registerclasses ( [ TOF_BTPARAMSOCS ] ) ;
end.

