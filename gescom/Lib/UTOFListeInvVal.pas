{***********UNITE*************************************************
Auteur  ...... : TG
Créé le ...... : 14/03/2000
Modifié le ... : 10/05/2000
Description .. : TOF Pour l'écran de lancement des états de liste
Suite ........ : d'inventaire valorisé
Suite ........ : ATTENTION : Utilise la tof ListeInvPrep comme ancetre de
Suite ........ : la tof ListeInvVal
Mots clefs ... : INVENTAIRE;VALORISE;STOCK
*****************************************************************}
unit UTOFListeInvVal;

interface
uses Classes, Controls, HCtrls, HEnt1, Hmsgbox, EntGC,         
{$IFDEF EAGLCLIENT}
     UTOB,MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
     UTofListeInvPrep;

function GCLanceFiche_ListeInvVal(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

type
    TOF_ListeInvVal = class(TOF_ListeInvPrep)
    private
      FSelectCombo : THValComboBox;
      EnContremarque : boolean;
    published
      procedure OnNew; override;
      procedure OnArgument(Arguments : String) ; override ;
      procedure OnUpdate ; override ;
      procedure OnLoad; override;
    end;

Const
	// libellés des messages
    TexteMessage: array[1..1] of string 	= (
          {1}        'Le code liste est inexistant'
            );

implementation
uses stdctrls;

function GCLanceFiche_ListeInvVal(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_ListeInvVal.OnNew;
begin
inherited;
FSelectCombo := THValComboBox(GetControl('KELPRIX'));
FSelectCombo.ItemIndex := 0;
end;

procedure TOF_ListeInvVal.OnArgument(Arguments: String);
var Ctl: TControl ;
begin
Inherited ;
EnContremarque := (Ecran.Name = 'GCLISTEINVVALCON');
if (ctxMode in V_PGI.PGIContexte) and (not EnContremarque) then
   begin
   SetControlProperty('GIL_DEPOT','Plus','GDE_SURSITE="X"');
   SetControlVisible('GIL_EMPLACEMENT', False);
   SetControlVisible('TGIL_EMPLACEMENT', False);
   SetControlVisible('CHKLOTS', False);
   SetControlChecked('CHKLOTS', False);
   end;

// Libellé Dépôt ou Etablissement
if not VH_GC.GCMultiDepots then
   begin
   if EnContremarque then Ctl := GetControl('TGIM_DEPOT')
   else Ctl := GetControl('TGIL_DEPOT') ;
    if (Ctl <> Nil) and (Ctl is THLabel) then THLabel(Ctl).Caption := 'Etablissement' ;
   end ;
end;

procedure TOF_ListeInvVal.OnUpdate ;
var QQ : TQuery;
begin
inherited ;
QQ := OpenSQL('SELECT GIE_CODELISTE FROM LISTEINVENT '+
              'WHERE GIE_CODELISTE="'+GetControlText('GIE_CODELISTE')+'"', true);
if QQ.EOF then PgiInfo (TexteMessage[1], Ecran.Caption);
Ferme(QQ) ;
end;

procedure TOF_ListeInvVal.OnLoad;
var QQ : TQuery ;
begin
inherited;
if EnContremarque then
   begin
   SetControlText('XX_VARIABLE10',RechDom('GCINFPRIXCON', FSelectCombo.Value, true));
   QQ:=OpenSql('SELECT CO_LIBRE FROM COMMUN WHERE CO_TYPE="GIM" AND CO_CODE="'+FSelectCombo.Value +'"', True);
   if not QQ.EOF then SetControlText('XX_VARIABLE11',QQ.Fields[0].AsString);
   end else
   begin
   SetControlText('XX_VARIABLE10',RechDom('GCINFPRIX', FSelectCombo.Value, true));
   QQ:=OpenSql('SELECT CO_LIBRE FROM COMMUN WHERE CO_TYPE="GIP" AND CO_CODE="'+FSelectCombo.Value +'"', True);
   if not QQ.EOF then SetControlText('XX_VARIABLE11',QQ.Fields[0].AsString);
   end;
Ferme(QQ) ;
if GetCheckBoxState('CHKDTL') <> cbChecked then  SetControlChecked('CHKLOTS',False);
if GetCheckBoxState('CHKECART') <> cbChecked then  SetControlChecked('CHKNONUL',False);
end;


initialization
RegisterClasses([TOF_ListeInvVal]);

end.
