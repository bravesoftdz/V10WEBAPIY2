unit UTofEtiqCli;

interface
uses UTOF,SysUtils,HCtrls,ComCtrls,
{$IFNDEF EAGLCLIENT}
     Mul,Fiche,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
     eMul,eFiche,MainEagl,
{$ENDIF}
     Classes,stdctrls,HEnt1,Ent1,EntGC,HQry,UTOB,forms
{$IFDEF GRC}
     ,UtilSelection,UtilRT
{$ENDIF}
     ;

Type
     TOF_EtiqCli = Class(TOF)
       private
         function Vire_Contact : boolean;
       public
       F : TForm;
       //procedure ChangementEtat;
       procedure OnLoad ; override;
       procedure OnUpdate ;  override;
       procedure OnChange(Sender: TObject);
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnClose ;  override;

     end;



Function RTLanceFiche_EtiqCli(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

Function RTLanceFiche_EtiqCli(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_EtiqCli.OnUpdate ;
var stWhere, Sql : string;
    QEtiq : TQuery;
    iInd, nbEtiq : integer;
begin
Inherited;
ExecuteSQL('DELETE FROM GCTMPETQCLI WHERE GZK_UTILISATEUR = "'+V_PGI.USer+'"');
SetControlText('XX_WHERE_USER','');
stWhere := RecupWhereCritere(TPageControl(GetControl('PAGES')));
SetControlText('XX_WHERE_USER','GZK_UTILISATEUR="'+V_PGI.USer+'"');

Sql:='SELECT T_TIERS FROM TIERS LEFT JOIN TIERSCOMPL ON LIEN(TIERS,TIERSCOMPL) ';
//JS 16/10/03 ça ne peut pas fonctionner en création de nouvel état !
//if copy(GetControlText('FETAT'),2,2)<>'C1' Then
if not Vire_Contact then
   Sql:=Sql+'left outer join CONTACT on C_AUXILIAIRE=T_AUXILIAIRE ';
{$IFDEF GRC}
if F.Name='RTETIQCLI' then
   Sql:=Sql+'LEFT JOIN PROSPECTS ON LIEN(TIERS,PROSPECTS) ';
{$ENDIF}
Sql:=Sql+stWhere;

QEtiq := OpenSQL(Sql, true);
While not QEtiq.Eof do
begin
  nbEtiq := StrToInt(GetControlText('NBEXEMPLAIRE'));
  for iInd := 0 to nbEtiq - 1 do ExecuteSQL('Insert into GCTMPETQCLI (GZK_UTILISATEUR, GZK_COMPTEUR, GZK_TIERS) ' +
    'values ("' + V_PGI.User + '",' + intToStr(iInd + 1) + ', "'+QEtiq.FindField('T_TIERS').AsString+'")');
  QEtiq.Next;
end;
Ferme(QEtiq);
end;

procedure TOF_EtiqCli.OnChange(Sender: TObject) ;
begin
if Vire_Contact Then
   begin
   THValComboBox(GetControl('C_NOM')).Text := '';
   THValComboBox(GetControl('C_FONCTION')).Value := '';
   TCheckBox(GetControl('C_PRINCIPAL')).State := cbGrayed;
   SetControlVisible('CONTACTS',FALSE) ;
   end
else
    begin
    SetControlVisible('CONTACTS',TRUE);
    SetFocusControl('CONTACTS');
    end;
end;

procedure TOF_EtiqCli.OnLoad ;
begin
inherited;
(*
if GetControlText('FETAT')='EC1' Then
   begin
   THValComboBox(GetControl('C_NOM')).Text := '';
   THValComboBox(GetControl('C_FONCTION')).Value := '';
   TCheckBox(GetControl('C_PRINCIPAL')).State := cbGrayed;
   end;
*)
{$IFDEF GRC}
  if F.Name='RTETIQCLI' then
     SetControlText('XX_WHERE',RTXXWhereConfident('CON')) ;
{$ENDIF}     
end;

procedure TOF_EtiqCli.OnArgument(Arguments : String ) ;
begin
inherited;
{$IFDEF GRC}
  F := TForm (Ecran);
  if F.Name='RTETIQCLI' then
    MulCreerPagesCL(F,'NOMFIC=GCTIERS');
{$ENDIF}
THValComboBox(GetControl('FEtat')).OnChange := OnChange;
SetControlText('XX_WHERE_USER','GZK_UTILISATEUR="'+V_PGI.USer+'"');
end;

procedure TOF_EtiqCli.OnClose ;
begin
Inherited;
ExecuteSQL('DELETE FROM GCTMPETQCLI WHERE GZK_UTILISATEUR = "'+V_PGI.USer+'"');
end;

function TOF_EtiqCli.Vire_Contact : boolean;
begin
 Result := not ((GetControlText('FETAT')='EC2') or (GetControlText('FETAT')='EGI') or (GetControlText('FETAT')='GEC'));
end;

Initialization
registerclasses([TOF_EtiqCli]) ;
end.
